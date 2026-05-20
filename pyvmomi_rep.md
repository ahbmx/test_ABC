```python
#!/usr/bin/env python3

"""
VMware vCenter inventory report using pyVmomi + pandas + xlsxwriter

Requirements:
    pip install pyvmomi pandas xlsxwriter openpyxl

What this script does:
    - Connects to multiple vCenters
    - Disables SSL verification
    - Uses common username/password for all vCenters
    - Collects:
        1. ESXi host WWNs
        2. VM disks + datastore mapping
        3. VM RDM mappings (NAA IDs, disk labels, compatibility mode)
    - Builds pandas DataFrames
    - Exports to Excel with:
        - Autofilter
        - Auto column width
        - Date in filename
"""

import ssl
from datetime import datetime

import pandas as pd
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import atexit


# =============================================================================
# CONFIG
# =============================================================================

VCENTERS = [
    "vcenter01.domain.local",
    "vcenter02.domain.local",
]

USERNAME = "administrator@vsphere.local"
PASSWORD = "SuperSecretPassword"


# =============================================================================
# SSL CONTEXT (DISABLE CERT VALIDATION)
# =============================================================================

ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
ssl_context.check_hostname = False
ssl_context.verify_mode = ssl.CERT_NONE


# =============================================================================
# HELPERS
# =============================================================================

def get_cluster_name(host):
    """Return cluster name for ESXi host."""
    try:
        parent = host.parent
        if isinstance(parent, vim.ClusterComputeResource):
            return parent.name
    except Exception:
        pass
    return "N/A"


def get_host_wwns(host):
    """
    Get FC HBA WWNs from ESXi host.
    """
    results = []

    try:
        storage_system = host.config.storageDevice

        for hba in storage_system.hostBusAdapter:
            if isinstance(hba, vim.host.FibreChannelHba):

                node_wwn = hba.nodeWorldWideName
                port_wwn = hba.portWorldWideName

                def format_wwn(wwn):
                    hexstr = f"{wwn:016x}"
                    return ":".join(
                        hexstr[i:i + 2] for i in range(0, 16, 2)
                    )

                results.append({
                    "Cluster": get_cluster_name(host),
                    "ESX Host": host.name,
                    "HBA Device": hba.device,
                    "Node WWN": format_wwn(node_wwn),
                    "Port WWN": format_wwn(port_wwn),
                })

    except Exception as e:
        print(f"Error collecting WWNs for host {host.name}: {e}")

    return results


def get_vm_disks(vm, cluster_name, esx_name, vcenter_name):
    """
    Get VM disk and datastore information.
    """
    results = []

    try:
        if not vm.config or not vm.config.hardware:
            return results

        for device in vm.config.hardware.device:

            if isinstance(device, vim.vm.device.VirtualDisk):

                datastore_name = "N/A"

                try:
                    if device.backing.datastore:
                        datastore_name = device.backing.datastore.name
                except Exception:
                    pass

                disk_label = device.deviceInfo.label
                capacity_gb = round(device.capacityInKB / 1024 / 1024, 2)

                results.append({
                    "vCenter": vcenter_name,
                    "Cluster": cluster_name,
                    "ESX Host": esx_name,
                    "VM": vm.name,
                    "Disk": disk_label,
                    "Capacity GB": capacity_gb,
                    "Datastore": datastore_name,
                })

    except Exception as e:
        print(f"Error collecting VM disks for {vm.name}: {e}")

    return results


def get_vm_rdms(vm, cluster_name, esx_name, vcenter_name):
    """
    Get VM RDM information.
    """
    results = []

    try:
        if not vm.config or not vm.config.hardware:
            return results

        for device in vm.config.hardware.device:

            if isinstance(device, vim.vm.device.VirtualDisk):

                backing = device.backing

                if isinstance(
                    backing,
                    vim.vm.device.VirtualDisk.RawDiskMappingVer1BackingInfo
                ):

                    naa_id = backing.deviceName
                    compatibility = backing.compatibilityMode
                    disk_label = device.deviceInfo.label

                    results.append({
                        "vCenter": vcenter_name,
                        "Cluster": cluster_name,
                        "ESX Host": esx_name,
                        "VM": vm.name,
                        "Disk": disk_label,
                        "NAA Device": naa_id,
                        "Compatibility Mode": compatibility,
                        "RDM File": backing.fileName,
                    })

    except Exception as e:
        print(f"Error collecting RDMs for {vm.name}: {e}")

    return results


def auto_adjust_columns(writer, df, sheet_name):
    """
    Auto adjust Excel column widths and apply autofilter.
    """
    worksheet = writer.sheets[sheet_name]

    # Autofilter
    worksheet.autofilter(
        0,
        0,
        len(df),
        len(df.columns) - 1
    )

    # Auto width
    for idx, col in enumerate(df.columns):

        max_len = max(
            (
                df[col].astype(str).map(len).max(),
                len(str(col))
            )
        ) + 2

        worksheet.set_column(idx, idx, max_len)


# =============================================================================
# MAIN
# =============================================================================

all_wwn_data = []
all_vm_disk_data = []
all_rdm_data = []

for vc in VCENTERS:

    print(f"\nConnecting to {vc} ...")

    try:
        si = SmartConnect(
            host=vc,
            user=USERNAME,
            pwd=PASSWORD,
            sslContext=ssl_context
        )

        atexit.register(Disconnect, si)

        content = si.RetrieveContent()

        container = content.viewManager.CreateContainerView(
            content.rootFolder,
            [vim.HostSystem, vim.VirtualMachine],
            True
        )

        objects = container.view

        # ---------------------------------------------------------------------
        # HOST WWNs
        # ---------------------------------------------------------------------
        hosts = [obj for obj in objects if isinstance(obj, vim.HostSystem)]

        for host in hosts:
            all_wwn_data.extend(get_host_wwns(host))

        # ---------------------------------------------------------------------
        # VM DATA
        # ---------------------------------------------------------------------
        vms = [obj for obj in objects if isinstance(obj, vim.VirtualMachine)]

        for vm in vms:

            try:
                runtime_host = vm.runtime.host

                if runtime_host:
                    esx_name = runtime_host.name
                    cluster_name = get_cluster_name(runtime_host)
                else:
                    esx_name = "N/A"
                    cluster_name = "N/A"

                all_vm_disk_data.extend(
                    get_vm_disks(
                        vm,
                        cluster_name,
                        esx_name,
                        vc
                    )
                )

                all_rdm_data.extend(
                    get_vm_rdms(
                        vm,
                        cluster_name,
                        esx_name,
                        vc
                    )
                )

            except Exception as e:
                print(f"Error processing VM {vm.name}: {e}")

        container.Destroy()

        print(f"Completed {vc}")

    except Exception as e:
        print(f"Failed to connect to {vc}: {e}")


# =============================================================================
# DATAFRAMES
# =============================================================================

df_wwns = pd.DataFrame(all_wwn_data)

df_vm_disks = pd.DataFrame(all_vm_disk_data)

df_rdms = pd.DataFrame(all_rdm_data)


# =============================================================================
# EXCEL EXPORT
# =============================================================================

date_str = datetime.now().strftime("%Y%m%d")
excel_file = f"vmware_inventory_report_{date_str}.xlsx"

with pd.ExcelWriter(
    excel_file,
    engine="xlsxwriter"
) as writer:

    # -------------------------------------------------------------------------
    # WWNs Sheet
    # -------------------------------------------------------------------------
    df_wwns.to_excel(
        writer,
        sheet_name="Host_WWNs",
        index=False
    )

    auto_adjust_columns(writer, df_wwns, "Host_WWNs")

    # -------------------------------------------------------------------------
    # VM Disks Sheet
    # -------------------------------------------------------------------------
    df_vm_disks.to_excel(
        writer,
        sheet_name="VM_Disks",
        index=False
    )

    auto_adjust_columns(writer, df_vm_disks, "VM_Disks")

    # -------------------------------------------------------------------------
    # RDMs Sheet
    # -------------------------------------------------------------------------
    df_rdms.to_excel(
        writer,
        sheet_name="VM_RDMs",
        index=False
    )

    auto_adjust_columns(writer, df_rdms, "VM_RDMs")

print(f"\nExcel report generated: {excel_file}")
```
