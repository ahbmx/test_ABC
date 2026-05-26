That’s a very good use case for Excel + VBA.
You can build a workbook that:

* Stores host/HBA/array information
* Automatically generates Brocade CLI commands
* Supports Virtual Fabrics (FID-aware)
* Separates Edge vs Core behavior
* Only shows commands when required fields are filled
* Exports ready-to-paste zoning scripts

A clean design is usually:

---

# Workbook Structure

## Sheet 1 — `INPUT`

Where you enter zoning data.

Example columns:

| A     | B          | C                       | D             | E                       | F              | G        | H      |
| ----- | ---------- | ----------------------- | ------------- | ----------------------- | -------------- | -------- | ------ |
| Host  | HBA Alias  | HBA WWPN                | Storage Alias | Storage WWPN            | Zone Name      | VSAN/FID | Fabric |
| app01 | APP01_HBA0 | 10:00:00:00:c9:aa:11:22 | PURE_CTL0     | 50:0a:09:81:00:aa:bb:cc | Z_APP01_PURE_A | 100      | CORE-A |

---

## Sheet 2 — `COMMANDS`

Auto-generated commands appear here.

Example:

```bash
fosconfig --fid 100
alicreate "APP01_HBA0","10:00:00:00:c9:aa:11:22"
alicreate "PURE_CTL0","50:0a:09:81:00:aa:bb:cc"
zonecreate "Z_APP01_PURE_A","APP01_HBA0;PURE_CTL0"
cfgadd "PROD_CFG","Z_APP01_PURE_A"
cfgenable "PROD_CFG"
cfgsave
```

---

# Recommended Logic

Since you use:

* Edge switches → hosts only
* Core switches → arrays + zoning

You should only generate:

* alias commands
* zoning commands
* cfgsave/cfgenable

for the CORE fabrics.

---

# Best Approach

## Use VBA Worksheet Events

You said:

> only display the commands when the cells are updated and not empty

Perfect use case for:

```vba
Worksheet_Change()
```

---

# VBA Example

In the `INPUT` sheet code:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)

    Call GenerateCommands

End Sub
```

---

# Main VBA Generator

Put this in a Module:

```vba
Sub GenerateCommands()

    Dim wsIn As Worksheet
    Dim wsCmd As Worksheet

    Dim lastRow As Long
    Dim i As Long
    Dim outRow As Long

    Set wsIn = Sheets("INPUT")
    Set wsCmd = Sheets("COMMANDS")

    wsCmd.Cells.Clear

    outRow = 1

    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        Dim hostAlias As String
        Dim hostWWPN As String
        Dim storageAlias As String
        Dim storageWWPN As String
        Dim zoneName As String
        Dim fid As String

        hostAlias = Trim(wsIn.Cells(i, 2).Value)
        hostWWPN = Trim(wsIn.Cells(i, 3).Value)

        storageAlias = Trim(wsIn.Cells(i, 4).Value)
        storageWWPN = Trim(wsIn.Cells(i, 5).Value)

        zoneName = Trim(wsIn.Cells(i, 6).Value)
        fid = Trim(wsIn.Cells(i, 7).Value)

        ' Only generate if mandatory fields exist
        If hostAlias <> "" And _
           hostWWPN <> "" And _
           storageAlias <> "" And _
           storageWWPN <> "" And _
           zoneName <> "" Then

            wsCmd.Cells(outRow, 1).Value = "setcontext " & fid
            outRow = outRow + 1

            wsCmd.Cells(outRow, 1).Value = _
                "alicreate """ & hostAlias & """,""" & hostWWPN & """"
            outRow = outRow + 1

            wsCmd.Cells(outRow, 1).Value = _
                "alicreate """ & storageAlias & """,""" & storageWWPN & """"
            outRow = outRow + 1

            wsCmd.Cells(outRow, 1).Value = _
                "zonecreate """ & zoneName & """,""" & _
                hostAlias & ";" & storageAlias & """"
            outRow = outRow + 1

            wsCmd.Cells(outRow, 1).Value = _
                "cfgadd ""PROD_CFG"",""" & zoneName & """"
            outRow = outRow + 2

        End If

    Next i

    wsCmd.Cells(outRow, 1).Value = "cfgsave"
    outRow = outRow + 1

    wsCmd.Cells(outRow, 1).Value = "cfgenable ""PROD_CFG"""

End Sub
```

---

# Recommended Improvements

## 1. Prevent Duplicate Aliases

Very important with Brocade.

Use a Dictionary object:

```vba
Dim dictAlias As Object
Set dictAlias = CreateObject("Scripting.Dictionary")
```

Then only generate alias once.

---

## 2. Multi-Initiator / Single-Initiator Support

You can add a column:

| Zoning Type |
| ----------- |
| SI          |
| MI          |

Then generate different zone memberships.

---

## 3. Edge/Core Separation

Add:

| Switch Type |
| ----------- |
| EDGE        |
| CORE        |

Then:

```vba
If switchType = "CORE" Then
```

Only generate zoning there.

---

## 4. Generate Different Scripts

Very useful tabs:

| Sheet       | Purpose                |
| ----------- | ---------------------- |
| CORE_A_CMDS | Fabric A               |
| CORE_B_CMDS | Fabric B               |
| EDGE_CMDS   | Optional aliases only  |
| ROLLBACK    | alidelete / zonedelete |
| AUDIT       | zoneobjectexp output   |

---

# Recommended Naming Convention

For X7/X8 fabrics:

## Aliases

```text
SRV_APP01_HBA0
SRV_APP01_HBA1
PURE01_CTL0P0
```

## Zones

```text
Z_APP01_PURE01_A
Z_APP01_PURE01_B
```

## Configs

```text
CFG_PROD_A
CFG_PROD_B
```

---

# Very Useful Features You Can Add

## Automatic Validation

Detect:

* Invalid WWPN format
* Duplicate aliases
* Duplicate zone names
* Wrong fabric pairing
* Wrong FID

---

## Button to Export TXT

```vba
Sub ExportCommands()

    Dim path As String

    path = ThisWorkbook.Path & "\core_a_zoning.txt"

    Open path For Output As #1

    Print #1, Sheets("COMMANDS").Range("A1").CurrentRegion.Value

    Close #1

End Sub
```

---

# Real Enterprise Enhancement

You can even generate:

* peer zones
* smart zoning
* LSAN configs
* MAPS commands
* port naming
* switchshow audit templates

from the same workbook.

---

# Architecture I’d Recommend

For your environment:

```text
INPUT
 ↓
VALIDATION
 ↓
VBA ENGINE
 ↓
CORE_A_SCRIPT
CORE_B_SCRIPT
ROLLBACK_SCRIPT
AUDIT_SCRIPT
```

That scales really well for large Brocade environments.

---

# Important X7/X8 Notes

For Gen7:

* `setcontext <FID>` is preferred in VF environments
* peer zoning is often better than standard zoning
* avoid duplicate aliases across logical switches
* use `cfgtransshow` awareness if multiple admins exist

---

# Best Practice

Generate commands only when:

```vba
If WorksheetFunction.CountA(Rows(i)) > 0 Then
```

AND mandatory fields are present.

This avoids partial/broken zoning commands.

---

If you want, I can also help you build:

1. A full enterprise-ready VBA workbook structure
2. Automatic duplicate checking
3. Peer zoning support
4. Fabric A/B automatic generation
5. Rollback script generation
6. Smart zoning support
7. A ready-to-import `.bas` VBA module
8. A polished Excel UI with buttons/forms
9. WWPN validation regex
10. Auto-generated zone naming based on hostname + storage type



Your code is generating commands row-by-row in sequence, so Excel outputs them as repeated “blocks”:

```text
setcontext
alicreate host
alicreate storage
zonecreate
cfgadd

setcontext
alicreate host
...
```

What you want instead is usually:

1. All `setcontext` grouped by FID
2. All `alicreate`
3. All `zonecreate`
4. All `cfgadd`
5. Then `cfgsave` / `cfgenable`

The cleanest way is to build each command type separately using string collections/dictionaries, then output them grouped.

Try this version:

```vb
Sub GenerateCommands()

    Dim wsIn As Worksheet
    Dim wsCmd As Worksheet

    Dim lastRow As Long
    Dim i As Long
    Dim outRow As Long

    Set wsIn = Sheets("INPUT")
    Set wsCmd = Sheets("COMMANDS")

    wsCmd.Cells.Clear

    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    ' Collections for grouped commands
    Dim ctxCmds As Collection
    Dim aliCmds As Collection
    Dim zoneCmds As Collection
    Dim cfgCmds As Collection

    Set ctxCmds = New Collection
    Set aliCmds = New Collection
    Set zoneCmds = New Collection
    Set cfgCmds = New Collection

    ' Prevent duplicate setcontext
    Dim dictFID As Object
    Set dictFID = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow

        Dim hostAlias As String
        Dim hostWWPN As String
        Dim storageAlias As String
        Dim storageWWPN As String
        Dim zoneName As String
        Dim fid As String

        hostAlias = Trim(wsIn.Cells(i, 2).Value)
        hostWWPN = Trim(wsIn.Cells(i, 3).Value)

        storageAlias = Trim(wsIn.Cells(i, 4).Value)
        storageWWPN = Trim(wsIn.Cells(i, 5).Value)

        zoneName = Trim(wsIn.Cells(i, 6).Value)
        fid = Trim(wsIn.Cells(i, 7).Value)

        ' Validate required fields
        If hostAlias <> "" And _
           hostWWPN <> "" And _
           storageAlias <> "" And _
           storageWWPN <> "" And _
           zoneName <> "" Then

            ' Add setcontext only once per FID
            If Not dictFID.Exists(fid) Then
                ctxCmds.Add "setcontext " & fid
                dictFID.Add fid, True
            End If

            ' Alias commands
            aliCmds.Add "alicreate """ & hostAlias & """,""" & hostWWPN & """"

            aliCmds.Add "alicreate """ & storageAlias & """,""" & storageWWPN & """"

            ' Zone commands
            zoneCmds.Add "zonecreate """ & zoneName & """,""" & _
                         hostAlias & ";" & storageAlias & """"

            ' CFG commands
            cfgCmds.Add "cfgadd ""PROD_CFG"",""" & zoneName & """"

        End If

    Next i

    ' Output grouped commands
    outRow = 1

    Dim item As Variant

    ' Context section
    For Each item In ctxCmds
        wsCmd.Cells(outRow, 1).Value = item
        outRow = outRow + 1
    Next item

    outRow = outRow + 1

    ' Alias section
    For Each item In aliCmds
        wsCmd.Cells(outRow, 1).Value = item
        outRow = outRow + 1
    Next item

    outRow = outRow + 1

    ' Zone section
    For Each item In zoneCmds
        wsCmd.Cells(outRow, 1).Value = item
        outRow = outRow + 1
    Next item

    outRow = outRow + 1

    ' CFGADD section
    For Each item In cfgCmds
        wsCmd.Cells(outRow, 1).Value = item
        outRow = outRow + 1
    Next item

    outRow = outRow + 1

    ' Final commands
    wsCmd.Cells(outRow, 1).Value = "cfgsave"
    outRow = outRow + 1

    wsCmd.Cells(outRow, 1).Value = "cfgenable ""PROD_CFG"""

    MsgBox "Commands generated successfully."

End Sub
```

This produces output like:

```text
setcontext 10

alicreate "HOST1","10:00:00:00"
alicreate "STOR1","20:00:00:00"
alicreate "HOST2","10:00:00:01"

zonecreate "ZONE1","HOST1;STOR1"
zonecreate "ZONE2","HOST2;STOR1"

cfgadd "PROD_CFG","ZONE1"
cfgadd "PROD_CFG","ZONE2"

cfgsave
cfgenable "PROD_CFG"
```

You can also improve it further by:

* removing duplicate `alicreate`
* grouping by FID automatically
* generating one cfg per fabric
* sorting commands alphabetically

if you want, I can help you build the “production-grade” version commonly used for Brocade zoning automation.
