Here is a comprehensive Excel VBA cheat sheet, organized by category for quick reference.

---

# 📘 EXCEL VBA CHEAT SHEET

## 1. BASIC SYNTAX

| Task | Code |
|------|------|
| Comment | `' This is a comment` |
| Line continuation | `Range("A1") = 10 & _`<br>`                " text"` |
| End statement | `Exit Sub`, `Exit Function`, `Exit For`, `Exit Do` |
| Stop execution | `Stop` |
| Wait | `Application.Wait Now + TimeValue("00:00:05")` |

## 2. VARIABLES & DATA TYPES

### Declaration
```vba
Dim varName As DataType
Public varName As DataType
Private varName As DataType
Static varName As DataType
```

### Common Data Types

| Type | Bytes | Range |
|------|-------|-------|
| `Integer` | 2 | -32,768 to 32,767 |
| `Long` | 4 | -2.1B to 2.1B |
| `Single` | 4 | ±3.4E38 |
| `Double` | 8 | ±1.8E308 |
| `Currency` | 8 | ±922T |
| `String` | variable | 0 to 2B chars |
| `Boolean` | 2 | True/False |
| `Date` | 8 | Jan 1 100 to Dec 31 9999 |
| `Variant` | varies | any type |

### Constants & Enums
```vba
Const PI As Double = 3.14159

Enum Days
    Monday = 1
    Tuesday = 2
End Enum
```

## 3. PROCEDURES

### Sub (no return value)
```vba
Sub MySub(Optional ByVal arg1 As String = "default")
    ' code
End Sub
```

### Function (returns value)
```vba
Function MyFunction(arg1 As Double) As Double
    MyFunction = arg1 * 2
End Function
```

### Call a procedure
```vba
Call MySub("hello")
MySub "hello"
Result = MyFunction(10)
```

## 4. CONDITIONAL STATEMENTS

### If...Then...Else
```vba
If condition Then
    ' code
ElseIf other Then
    ' code
Else
    ' code
End If

' One-liner
If x = 5 Then MsgBox "Five"
```

### Select Case
```vba
Select Case var
    Case 1
        ' code
    Case 2 To 5
        ' code
    Case Is > 10
        ' code
    Case "apple", "orange"
        ' code
    Case Else
        ' code
End Select
```

## 5. LOOPS

### For...Next
```vba
For i = 1 To 10 Step 2
    ' code
Next i

' For Each
Dim cell As Range
For Each cell In Range("A1:A10")
    cell.Value = 0
Next cell
```

### Do While / Until
```vba
Do While condition
    ' code
Loop

Do Until condition
    ' code
Loop

Do
    ' code
Loop While condition

Do
    ' code
Loop Until condition
```

### While...Wend
```vba
While condition
    ' code
Wend
```

## 6. RANGE & CELLS REFERENCING

```vba
Range("A1") = 10
Cells(1, 1) = 10           ' Row 1, Col 1
Range("A1:B5").Value = 100
Range("A:A")               ' Entire column
Range("1:1")               ' Entire row

' Offset
Range("A1").Offset(1, 0) = 20   ' One row down

' Resize
Range("A1").Resize(3, 2) = 50   ' 3x2 area

' Named range
Range("MyName").Value = "Test"

' Current region
Range("A1").CurrentRegion

' End (Ctrl+Arrow)
Range("A1").End(xlDown)
' xlUp, xlToLeft, xlToRight

' Special cells
Range("A1:A10").SpecialCells(xlCellTypeBlanks)

' Entire row/column
Range("B2").EntireRow.Delete
```

### Properties

| Property | Example |
|----------|---------|
| Value | `Range("A1").Value = 5` |
| Formula | `Range("A1").Formula = "=B1+C1"` |
| Text | `MsgBox Range("A1").Text` |
| Address | `Range("A1").Address` |
| Count | `Range("A1:A10").Count` |
| Font | `Range("A1").Font.Bold = True` |
| Color | `Range("A1").Interior.Color = RGB(255,0,0)` |
| NumberFormat | `Range("A1").NumberFormat = "$#,##0.00"` |

## 7. WORKBOOKS & WORKSHEETS

### Workbooks
```vba
ThisWorkbook                    ' Workbook containing code
ActiveWorkbook                  ' Currently active workbook
Workbooks("MyFile.xlsx")        ' By name
Workbooks(1)                    ' By index

' Open/Close/Save
Workbooks.Open "C:\path\file.xlsx"
ActiveWorkbook.Close SaveChanges:=False
ActiveWorkbook.Save
ActiveWorkbook.SaveAs "C:\newpath\file.xlsx"

' Add/Delete
Workbooks.Add
ActiveWorkbook.Close
```

### Worksheets
```vba
ActiveSheet
Worksheets("Sheet1")
Sheets("Sheet1")                ' Sheets includes charts
Sheet1                         ' Code name (best practice)

' Add/Delete/Move
Worksheets.Add Before:=Worksheets(1)
Worksheets("Sheet2").Delete
Worksheets("Sheet1").Move After:=Worksheets(3)

' Copy
Worksheets("Sheet1").Copy After:=Worksheets(3)
```

## 8. VARIABLES & ARRAYS

### Standard
```vba
Dim arr(1 To 5) As Integer      ' Fixed size
Dim arr() As Variant            ' Dynamic
ReDim arr(1 To 10)              ' Re-dimension
ReDim Preserve arr(1 To 20)     ' Preserve data
```

### Array operations
```vba
' Create from range
arr = Range("A1:A10").Value

' Write to range
Range("B1:B10").Value = arr

' Split/Join
arr = Split("a,b,c", ",")
result = Join(arr, "-")

' Functions
UBound(arr)                     ' Upper bound
LBound(arr)                     ' Lower bound (usually 1 or 0)
Erase arr                       ' Clear array
```

## 9. ERROR HANDLING

```vba
On Error GoTo ErrorHandler      ' Jump to label
On Error Resume Next            ' Ignore errors
On Error GoTo 0                 ' Disable error handling

' Example
Sub SafeDivide()
    On Error GoTo ErrHandler
    Dim result As Double
    result = 10 / 0
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
```

### Error Objects
```vba
Err.Number         ' Error code
Err.Description    ' Error message
Err.Source         ' Source of error
Err.Clear          ' Clear error
```

## 10. USER INTERACTION

### Message Box
```vba
MsgBox "Simple message"

response = MsgBox("Continue?", vbYesNo + vbQuestion, "Title")
' Returns: vbYes, vbNo, vbOK, vbCancel, vbAbort, vbRetry, vbIgnore

MsgBox "Alert", vbCritical
' Buttons: vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
' Icons: vbCritical, vbQuestion, vbExclamation, vbInformation
```

### Input Box
```vba
result = InputBox("Enter value", "Title", "default")
If result = "" Then Exit Sub
```

### Application InputBox (more control)
```vba
Set rng = Application.InputBox("Select range", Type:=8)
' Type: 0=formula, 1=number, 2=text, 4=boolean, 8=range, 16=error
```

## 11. CELL FORMATTING

```vba
' Number format
Range("A1").NumberFormat = "General"
Range("A1").NumberFormat = "0.00"
Range("A1").NumberFormat = "mm/dd/yyyy"
Range("A1").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

' Font
With Range("A1").Font
    .Name = "Calibri"
    .Size = 12
    .Bold = True
    .Italic = False
    .Underline = xlUnderlineStyleSingle
    .Color = RGB(0, 0, 255)
End With

' Interior (Background)
Range("A1").Interior.Color = RGB(255, 255, 0)
Range("A1").Interior.ColorIndex = 6
Range("A1").Interior.Pattern = xlSolid

' Borders
With Range("A1").Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .Color = RGB(0, 0, 0)
End With

' HorizontalAlignment: xlLeft, xlCenter, xlRight
' VerticalAlignment: xlTop, xlCenter, xlBottom
```

## 12. USEFUL METHODS

### Find
```vba
Set found = Range("A1:A100").Find(What:="search", LookAt:=xlPart)
If Not found Is Nothing Then
    found.Select
End If
```

### AutoFilter
```vba
Range("A1:C100").AutoFilter Field:=1, Criteria1:=">10"
Range("A1:C100").AutoFilter Field:=2, Criteria1:=Array("A","B"), Operator:=xlFilterValues
ActiveSheet.AutoFilterMode = False   ' Remove filter
```

### Sort
```vba
With ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("A1"), Order:=xlAscending
    .SetRange Range("A1:C100")
    .Header = xlYes
    .Apply
End With
```

### Special operations
```vba
Range("A1").Copy Destination:=Range("B1")
Range("A1").Cut Destination:=Range("B1")
Range("A1:A5").ClearContents
Range("A1:A5").Delete Shift:=xlUp
Range("A1:A5").Insert Shift:=xlDown
```

## 13. WORKBOOK EVENTS (in ThisWorkbook)

```vba
Private Sub Workbook_Open()
    ' Runs when workbook opens
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Runs before workbook closes
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Runs when any sheet changes
End Sub
```

## 14. WORKSHEET EVENTS (in sheet module)

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        MsgBox "A1 changed"
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Target.Interior.Color = RGB(200, 200, 255)
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' Cancel = True prevents default double-click
End Sub
```

## 15. USEFUL APPLICATION SETTINGS

```vba
Application.ScreenUpdating = False   ' Faster code
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False    ' Suppress warnings
Application.EnableEvents = False     ' Disable events
Application.StatusBar = "Processing..."
Application.CutCopyMode = False      ' Clear clipboard
Application.Wait Now + #12:00:05 AM#

' Reset to defaults
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.StatusBar = False
```

## 16. FILE & FOLDER OPERATIONS

```vba
' Check if file exists
Dir("C:\file.xlsx") <> ""

' File system object (requires reference to Microsoft Scripting Runtime)
Dim fso As New FileSystemObject
fso.FileExists("C:\file.txt")
fso.FolderExists("C:\folder")
fso.CreateFolder("C:\newfolder")
fso.DeleteFile "C:\file.txt"
fso.DeleteFolder "C:\folder"

' Kill (simple delete)
Kill "C:\file.txt"
```

## 17. OFTEN-USED CONSTANTS

| Constant | Value | Description |
|----------|-------|-------------|
| xlUp | -4162 | Move up |
| xlDown | -4121 | Move down |
| xlLeft | -4159 | Move left |
| xlRight | -4161 | Move right |
| xlPart | 2 | Partial match |
| xlWhole | 1 | Whole match |
| xlByRows | 1 | Sort by rows |
| xlByColumns | 2 | Sort by columns |
| xlYes | 1 | Yes |
| xlNo | 2 | No |
| xlContinuous | 1 | Solid line |
| xlDash | -4115 | Dashed line |
| xlDot | -4118 | Dotted line |

## 18. QUICK REFERENCE PATTERNS

### Loop through all used cells in a column
```vba
Dim i As Long
Dim lastRow As Long
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
For i = 1 To lastRow
    Debug.Print Cells(i, "A").Value
Next i
```

### Loop through all worksheets
```vba
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Range("A1").Value = "Updated"
Next ws
```

### Create a new workbook and copy data
```vba
Dim newWB As Workbook
Set newWB = Workbooks.Add
ThisWorkbook.Worksheets("Data").UsedRange.Copy newWB.Worksheets(1).Range("A1")
newWB.SaveAs "C:\Output.xlsx"
newWB.Close
```

### Disable events while running (recommended pattern)
```vba
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Your code here

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
```

---

This cheat sheet covers the most common VBA operations. Keep it bookmarked or print it for daily reference.
