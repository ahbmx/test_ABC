In Excel VBA, the standard way to automatically update sheets when cells change is by using **Worksheet Events**, especially:

* `Worksheet_Change`
* `Worksheet_Calculate`
* `Workbook_SheetChange`

---

# 1. Update Sheet When a Cell Changes

## Basic `Worksheet_Change` Event

This runs automatically whenever a user edits a cell.

### Example

Put this inside the worksheet module (not a normal module):

1. Press `ALT + F11`
2. Double-click the sheet (example: `Sheet1`)
3. Paste:

```vba id="0o4p5v"
Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Target, Range("A1")) Is Nothing Then
        
        Range("B1").Value = "Updated: " & Now
        
    End If

End Sub
```

### What it does

* Watches `A1`
* When `A1` changes:

  * `B1` gets updated with timestamp

---

# 2. Update Entire Sheet Based on Changes

## Example: Recalculate Summary

```vba id="g7rv3x"
Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        
        Range("C1").Value = _
            Application.WorksheetFunction.Sum(Range("A:A"))
            
    End If

End Sub
```

### Result

Whenever anything in Column A changes:

* `C1` updates automatically

---

# 3. Trigger Only for Specific Range

```vba id="0f6f67"
Private Sub Worksheet_Change(ByVal Target As Range)

    If Intersect(Target, Range("A1:A10")) Is Nothing Then Exit Sub

    MsgBox "Tracked cells changed"

End Sub
```

---

# 4. Update Another Worksheet Automatically

## Example

```vba id="3h7v4o"
Private Sub Worksheet_Change(ByVal Target As Range)

    Worksheets("Report").Range("A1").Value = _
        Worksheets("Data").Range("A1").Value

End Sub
```

---

# 5. Prevent Infinite Loops (Very Important)

If your VBA changes cells inside `Worksheet_Change`, it can trigger itself forever.

Use:

```vba id="z6kn24"
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit

    Application.EnableEvents = False

    Range("B1").Value = UCase(Range("A1").Value)

SafeExit:

    Application.EnableEvents = True

End Sub
```

This is the professional pattern.

---

# 6. Watch Formula Changes

`Worksheet_Change` does NOT detect formula recalculations.

Use `Worksheet_Calculate` instead.

## Example

```vba id="u7iqta"
Private Sub Worksheet_Calculate()

    MsgBox "Sheet recalculated"

End Sub
```

Useful when:

* formulas update
* external links refresh
* pivot tables refresh

---

# 7. Workbook-Wide Cell Change Monitoring

Place this in `ThisWorkbook`:

```vba id="h9oc00"
Private Sub Workbook_SheetChange( _
    ByVal Sh As Object, _
    ByVal Target As Range)

    MsgBox "Changed on sheet: " & Sh.Name

End Sub
```

This tracks ALL sheets.

---

# 8. Real-World Example: Auto Timestamp

```vba id="1yfxca"
Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Target, Range("A:A")) Is Nothing Then

        Cells(Target.Row, 2).Value = Now

    End If

End Sub
```

### Result

When Column A changes:

* Column B gets timestamp

---

# 9. Real-World Example: Auto Uppercase

```vba id="l7ic5j"
Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Column = 1 Then

        Application.EnableEvents = False

        Target.Value = UCase(Target.Value)

        Application.EnableEvents = True

    End If

End Sub
```

---

# 10. Real-World Example: Sync Two Sheets

```vba id="muhj5v"
Private Sub Worksheet_Change(ByVal Target As Range)

    Worksheets("Backup").Range(Target.Address).Value = Target.Value

End Sub
```

---

# 11. Best Practice Template

```vba id="m7jlwm"
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Cleanup

    Application.EnableEvents = False

    ' Check changed range
    If Intersect(Target, Range("A1:A100")) Is Nothing Then GoTo Cleanup

    ' Your logic here
    Range("B1").Value = "Updated"

Cleanup:

    Application.EnableEvents = True

End Sub
```

---

# 12. Common Event Types

| Event                     | Trigger                |
| ------------------------- | ---------------------- |
| Worksheet_Change          | User edits cells       |
| Worksheet_Calculate       | Formula recalculates   |
| Worksheet_SelectionChange | User changes selection |
| Workbook_Open             | Workbook opens         |
| Workbook_BeforeClose      | Workbook closes        |

---

# 13. Where to Put the Code

| Code Type        | Location        |
| ---------------- | --------------- |
| Worksheet events | Sheet module    |
| Workbook events  | ThisWorkbook    |
| Normal macros    | Standard module |

---

# 14. Fastest Setup Method

## Create Event Automatically

In VBA Editor:

1. Open sheet module
2. Top-left dropdown → `Worksheet`
3. Top-right dropdown → `Change`

Excel generates:

```vba id="1m7khw"
Private Sub Worksheet_Change(ByVal Target As Range)

End Sub
```

---

# 15. Common Mistakes

## WRONG: Infinite Loop

```vba id="99e4wz"
Private Sub Worksheet_Change(ByVal Target As Range)

    Range("A1").Value = "Test"

End Sub
```

## RIGHT

```vba id="2j8v81"
Application.EnableEvents = False
Range("A1").Value = "Test"
Application.EnableEvents = True
```

---

# 16. Advanced Pattern — Detect Which Cell Changed

```vba id="88l4z6"
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim cell As Range

    For Each cell In Target

        Debug.Print cell.Address
        Debug.Print cell.Value

    Next cell

End Sub
```

Useful for:

* multi-cell paste
* logging
* audits

---

# 17. Example: Automatic Dashboard Refresh

```vba id="jw0j6m"
Private Sub Worksheet_Change(ByVal Target As Range)

    ThisWorkbook.RefreshAll

End Sub
```

---

# 18. Recommended Architecture for Large Files

For professional workbooks:

* Use worksheet events only for triggers
* Put business logic in standard modules
* Keep events lightweight

Example:

```vba id="rqq3p6"
Private Sub Worksheet_Change(ByVal Target As Range)

    Call UpdateDashboard

End Sub
```

Then:

```vba id="m3g6t7"
Sub UpdateDashboard()

    ' Main logic here

End Sub
```

This makes maintenance much easier.
