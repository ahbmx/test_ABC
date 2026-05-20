```python
#!/usr/bin/env python3

"""
VMware vCenter inventory report using pyVmomi + pandas + xlsxwriter
WITH progress bars.

Requirements:
    pip install pyvmomi pandas xlsxwriter tqdm

Features:
    - Multiple vCenters
    - SSL verification disabled
    - Common credentials
    - Progress bars for:
        * vCenters
        * Hosts
        * VMs
    - Excel report with:
        * Autofilter
        * Auto column width
        * Date-based filename
"""

import ssl
import atexit
from datetime import datetime

import pandas as pd
from tqdm import tqdm

from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim


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
# SSL CONTEXT
# =============================================================================

ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
ssl_context.check_hostname = False
ssl_context.verify_mode = ssl.CERT_NONE


# =============================================================================
# HELPERS
# =============================================================================

def get_cluster_name(host):
    """
    Return cluster name for ESXi host.
    """
    try:
        parent = host.parent
        if isinstance(parent, vim.ClusterComputeResource):
            return parent.name
    except Exception:
        pass

    return "N/A"


def format_wwn(wwn):
    """
    Convert integer WWN to readable format.
    """
    hexstr = f"{wwn:016x}"

    return ":".join(
        hexstr[i:i + 2]
        for i in range(0, 16, 2)
    )


def get_host_wwns(host):
    """
    Get FC HBA WWNs from ESXi host.
    """
    results = []

    try:
        storage_system = host.config.storageDevice

        for hba in storage_system.hostBusAdapter:

            if isinstance(hba, vim.host.FibreChannelHba):

                results.append({
                    "Cluster": get_cluster_name(host),
                    "ESX Host": host.name,
                    "HBA Device": hba.device,
                    "Node WWN": format_wwn(hba.nodeWorldWideName),
                    "Port WWN": format_wwn(hba.portWorldWideName),
                })

    except Exception as e:
        print(f"Error collecting WWNs from {host.name}: {e}")

    return results


def get_vm_disks(vm, cluster_name, esx_name, vcenter_name):
    """
    Get VM disks and datastore mappings.
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

                results.append({
                    "vCenter": vcenter_name,
                    "Cluster": cluster_name,
                    "ESX Host": esx_name,
                    "VM": vm.name,
                    "Disk": device.deviceInfo.label,
                    "Capacity GB": round(
                        device.capacityInKB / 1024 / 1024,
                        2
                    ),
                    "Datastore": datastore_name,
                })

    except Exception as e:
        print(f"Error collecting disks from {vm.name}: {e}")

    return results


def get_vm_rdms(vm, cluster_name, esx_name, vcenter_name):
    """
    Get VM RDM mappings.
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

                    results.append({
                        "vCenter": vcenter_name,
                        "Cluster": cluster_name,
                        "ESX Host": esx_name,
                        "VM": vm.name,
                        "Disk": device.deviceInfo.label,
                        "NAA Device": backing.deviceName,
                        "Compatibility Mode": backing.compatibilityMode,
                        "RDM File": backing.fileName,
                    })

    except Exception as e:
        print(f"Error collecting RDMs from {vm.name}: {e}")

    return results


def auto_adjust_columns(writer, df, sheet_name):
    """
    Apply autofilter and auto-size columns.
    """
    worksheet = writer.sheets[sheet_name]

    # Autofilter
    worksheet.autofilter(
        0,
        0,
        len(df),
        len(df.columns) - 1
    )

    # Auto column sizing
    for idx, col in enumerate(df.columns):

        max_len = max(
            df[col].astype(str).map(len).max(),
            len(col)
        ) + 2

        worksheet.set_column(idx, idx, max_len)


# =============================================================================
# MAIN
# =============================================================================

all_wwn_data = []
all_vm_disk_data = []
all_rdm_data = []

for vc in tqdm(VCENTERS, desc="vCenters", colour="green"):

    try:
        si = SmartConnect(
            host=vc,
            user=USERNAME,
            pwd=PASSWORD,
            sslContext=ssl_context
        )

        atexit.register(Disconnect, si)

        content = si.RetrieveContent()

        # ---------------------------------------------------------------------
        # HOSTS
        # ---------------------------------------------------------------------
        host_view = content.viewManager.CreateContainerView(
            content.rootFolder,
            [vim.HostSystem],
            True
        )

        hosts = host_view.view

        for host in tqdm(
            hosts,
            desc=f"{vc} Hosts",
            leave=False,
            colour="cyan"
        ):
            all_wwn_data.extend(get_host_wwns(host))

        host_view.Destroy()

        # ---------------------------------------------------------------------
        # VMS
        # ---------------------------------------------------------------------
        vm_view = content.viewManager.CreateContainerView(
            content.rootFolder,
            [vim.VirtualMachine],
            True
        )

        vms = vm_view.view

        for vm in tqdm(
            vms,
            desc=f"{vc} VMs",
            leave=False,
            colour="yellow"
        ):

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

        vm_view.Destroy()

    except Exception as e:
        print(f"Failed connecting to {vc}: {e}")


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

print("\nGenerating Excel report...")

with pd.ExcelWriter(
    excel_file,
    engine="xlsxwriter"
) as writer:

    # WWNs
    df_wwns.to_excel(
        writer,
        sheet_name="Host_WWNs",
        index=False
    )

    auto_adjust_columns(
        writer,
        df_wwns,
        "Host_WWNs"
    )

    # VM Disks
    df_vm_disks.to_excel(
        writer,
        sheet_name="VM_Disks",
        index=False
    )

    auto_adjust_columns(
        writer,
        df_vm_disks,
        "VM_Disks"
    )

    # RDMs
    df_rdms.to_excel(
        writer,
        sheet_name="VM_RDMs",
        index=False
    )

    auto_adjust_columns(
        writer,
        df_rdms,
        "VM_RDMs"
    )

print(f"\nReport generated successfully:\n{excel_file}")
```
