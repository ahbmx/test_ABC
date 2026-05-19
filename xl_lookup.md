You can look up values in Excel tables with VBA in a few common ways, depending on what you need:

* Use Excel worksheet functions (`VLOOKUP`, `XLOOKUP`, `MATCH/INDEX`)
* Search a `ListObject` (Excel Table)
* Loop through rows manually
* Use dictionaries for fast lookups

Here are the most useful approaches.

---

# 1. Use `XLOOKUP` from VBA (Excel 365 / newer Excel)

If your Excel version supports `XLOOKUP`, this is the cleanest method.

Example table:

| ID   | Name  |
| ---- | ----- |
| 1001 | Alice |
| 1002 | Bob   |

VBA:

```vb
Sub LookupWithXLookup()

    Dim result As Variant

    result = Application.WorksheetFunction.XLookup( _
                1002, _
                Range("A2:A10"), _
                Range("B2:B10") _
             )

    MsgBox result

End Sub
```

This returns `"Bob"`.

---

# 2. Use `VLookup` from VBA

Works in older Excel versions.

```vb
Sub LookupWithVLookup()

    Dim result As Variant

    result = Application.VLookup(1002, Range("A2:B10"), 2, False)

    MsgBox result

End Sub
```

Parameters:

* `1002` → value to find
* `A2:B10` → lookup table
* `2` → return value from 2nd column
* `False` → exact match

---

# 3. Lookup Inside an Excel Table (`ListObject`)

If your data is a real Excel Table (Insert → Table), this is usually best.

Example table name: `Employees`

```vb
Sub LookupInTable()

    Dim tbl As ListObject
    Dim foundCell As Range

    Set tbl = Sheet1.ListObjects("Employees")

    Set foundCell = tbl.ListColumns("ID").DataBodyRange.Find(1002)

    If Not foundCell Is Nothing Then
        MsgBox foundCell.Offset(0, 1).Value
    Else
        MsgBox "Not found"
    End If

End Sub
```

This:

* searches the `ID` column
* returns the value from the next column

---

# 4. Use `MATCH` + `INDEX` Style Logic

Very flexible.

```vb
Sub MatchExample()

    Dim rowNum As Variant
    Dim result As Variant

    rowNum = Application.Match(1002, Range("A2:A10"), 0)

    If Not IsError(rowNum) Then
        result = Range("B2:B10").Cells(rowNum).Value
        MsgBox result
    Else
        MsgBox "Not found"
    End If

End Sub
```

---

# 5. Fastest Method for Large Data: Dictionary Lookup

For thousands of rows, dictionaries are much faster.

```vb
Sub DictionaryLookup()

    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = Sheet1

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        dict(ws.Cells(i, 1).Value) = ws.Cells(i, 2).Value
    Next i

    MsgBox dict(1002)

End Sub
```

Good for:

* repeated lookups
* large datasets
* automation scripts

---

# Recommended Approach

| Situation      | Best Option         |
| -------------- | ------------------- |
| Modern Excel   | `XLOOKUP`           |
| Older Excel    | `VLOOKUP`           |
| Excel Table    | `ListObject + Find` |
| Large datasets | `Dictionary`        |
| Flexible logic | `MATCH`             |

---

# Practical Example: Lookup by Header Name

This is useful when column positions may change.

```vb
Sub DynamicColumnLookup()

    Dim tbl As ListObject
    Dim idCol As Range
    Dim nameCol As Range
    Dim foundCell As Range

    Set tbl = Sheet1.ListObjects("Employees")

    Set idCol = tbl.ListColumns("ID").DataBodyRange
    Set nameCol = tbl.ListColumns("Name").DataBodyRange

    Set foundCell = idCol.Find(1002)

    If Not foundCell Is Nothing Then
        MsgBox nameCol.Cells(foundCell.Row - idCol.Row + 1).Value
    End If

End Sub
```

This survives column reordering.

---

# Common Errors

### `Run-time error 1004`

Usually caused by:

* lookup value not found
* wrong range size
* bad table name

Handle safely:

```vb
If IsError(result) Then
    MsgBox "Not found"
End If
```

---

# Microsoft VBA References

* [Excel VBA documentation](https://learn.microsoft.com/en-us/office/vba/api/overview/excel?utm_source=chatgpt.com)
* [ListObject object reference](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject?utm_source=chatgpt.com)
* [WorksheetFunction.VLookup](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.vlookup?utm_source=chatgpt.com)
* [WorksheetFunction.XLookup](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.xlookup?utm_source=chatgpt.com)
