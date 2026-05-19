# Excel VBA Cheat Sheet

## What VBA Is

**VBA (Visual Basic for Applications)** is the programming language built into Microsoft Excel used to automate tasks, build tools, manipulate worksheets, create forms, and integrate with other Office apps.

---

# 1. VBA Editor Basics

## Open VBA Editor

| Action           | Shortcut    |
| ---------------- | ----------- |
| Open VBA Editor  | `ALT + F11` |
| Immediate Window | `CTRL + G`  |
| Object Browser   | `F2`        |
| Run Macro        | `F5`        |
| Step Into        | `F8`        |

---

# 2. VBA Project Structure

## Common Components

| Component    | Purpose                |
| ------------ | ---------------------- |
| Module       | Store macros/functions |
| UserForm     | Custom GUI             |
| Sheet Module | Worksheet events       |
| ThisWorkbook | Workbook events        |
| Class Module | Custom objects         |

---

# 3. Your First Macro

```vba
Sub HelloWorld()

    MsgBox "Hello World!"

End Sub
```

---

# 4. VBA Syntax Basics

## Variables

```vba
Dim x As Integer
Dim name As String
Dim price As Double
Dim isValid As Boolean
```

## Constants

```vba
Const PI As Double = 3.14159
```

## Comments

```vba
' This is a comment
```

---

# 5. Data Types

| Type    | Description         |
| ------- | ------------------- |
| Integer | Whole numbers       |
| Long    | Large whole numbers |
| Double  | Decimal numbers     |
| String  | Text                |
| Boolean | True/False          |
| Date    | Dates/times         |
| Variant | Any type            |
| Object  | Object references   |

---

# 6. Operators

## Arithmetic

| Operator | Meaning   |
| -------- | --------- |
| +        | Add       |
| -        | Subtract  |
| *        | Multiply  |
| /        | Divide    |
| ^        | Power     |
| Mod      | Remainder |

## Comparison

| Operator | Meaning       |
| -------- | ------------- |
| =        | Equal         |
| <>       | Not equal     |
| >        | Greater than  |
| <        | Less than     |
| >=       | Greater/equal |
| <=       | Less/equal    |

## Logical

| Operator | Meaning     |
| -------- | ----------- |
| And      | Both true   |
| Or       | Either true |
| Not      | Negation    |

---

# 7. Conditional Statements

## If Statement

```vba
If x > 10 Then
    MsgBox "Large"
Else
    MsgBox "Small"
End If
```

## ElseIf

```vba
If score >= 90 Then
    grade = "A"
ElseIf score >= 80 Then
    grade = "B"
Else
    grade = "C"
End If
```

## Select Case

```vba
Select Case dayNum

    Case 1
        MsgBox "Monday"

    Case 2
        MsgBox "Tuesday"

    Case Else
        MsgBox "Other"

End Select
```

---

# 8. Loops

## For Loop

```vba
Dim i As Integer

For i = 1 To 10
    Debug.Print i
Next i
```

## Step

```vba
For i = 10 To 1 Step -1
    Debug.Print i
Next i
```

## For Each

```vba
Dim cell As Range

For Each cell In Range("A1:A10")
    Debug.Print cell.Value
Next cell
```

## Do While

```vba
Do While x < 10
    x = x + 1
Loop
```

## Do Until

```vba
Do Until x = 10
    x = x + 1
Loop
```

---

# 9. Procedures

## Sub Procedure

```vba
Sub MyMacro()

End Sub
```

## Function Procedure

```vba
Function AddNumbers(a As Double, b As Double) As Double

    AddNumbers = a + b

End Function
```

Usage in worksheet:

```excel
=AddNumbers(5,3)
```

---

# 10. Working with Worksheets

## Reference Workbook

```vba
Workbooks("Sales.xlsx")
```

## Reference Worksheet

```vba
Worksheets("Sheet1")
```

## Activate Sheet

```vba
Sheets("Data").Activate
```

## Add Worksheet

```vba
Worksheets.Add
```

## Delete Worksheet

```vba
Application.DisplayAlerts = False
Sheets("Temp").Delete
Application.DisplayAlerts = True
```

---

# 11. Working with Cells

## Read Cell

```vba
value = Range("A1").Value
```

## Write Cell

```vba
Range("A1").Value = "Hello"
```

## Using Cells()

```vba
Cells(1, 1).Value = "Hello"
```

## Resize

```vba
Range("A1").Resize(5, 3)
```

## Offset

```vba
Range("A1").Offset(1, 0)
```

---

# 12. Range Operations

## Select Range

```vba
Range("A1:C10").Select
```

## Copy/Paste

```vba
Range("A1:A10").Copy Destination:=Range("B1")
```

## Clear

```vba
Range("A1").Clear
Range("A1").ClearContents
Range("A1").ClearFormats
```

## Merge Cells

```vba
Range("A1:C1").Merge
```

---

# 13. Find Last Row / Column

## Last Row

```vba
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
```

## Last Column

```vba
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
```

---

# 14. Formatting

## Font

```vba
Range("A1").Font.Bold = True
Range("A1").Font.Size = 14
```

## Interior Color

```vba
Range("A1").Interior.Color = vbYellow
```

## Number Format

```vba
Range("A1").NumberFormat = "$#,##0.00"
```

## Borders

```vba
Range("A1").Borders.LineStyle = xlContinuous
```

---

# 15. Message Boxes & Input

## MsgBox

```vba
MsgBox "Done!"
```

## Yes/No

```vba
answer = MsgBox("Continue?", vbYesNo)

If answer = vbYes Then
    MsgBox "Proceeding"
End If
```

## InputBox

```vba
name = InputBox("Enter name")
```

---

# 16. Arrays

## Static Array

```vba
Dim arr(1 To 5) As Integer
```

## Dynamic Array

```vba
Dim arr() As Integer
ReDim arr(1 To 10)
```

## Loop Through Array

```vba
Dim i As Integer

For i = LBound(arr) To UBound(arr)
    Debug.Print arr(i)
Next i
```

---

# 17. Collections & Dictionaries

## Collection

```vba
Dim col As New Collection

col.Add "Apple"
col.Add "Banana"
```

## Dictionary

Requires:
`Microsoft Scripting Runtime`

```vba
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

dict.Add "A", 100

Debug.Print dict("A")
```

---

# 18. Error Handling

## Basic Error Handler

```vba
On Error GoTo ErrorHandler

x = 1 / 0

Exit Sub

ErrorHandler:
    MsgBox Err.Description
```

## Ignore Errors

```vba
On Error Resume Next
```

## Reset Errors

```vba
On Error GoTo 0
```

---

# 19. Debugging

## Immediate Window

```vba
Debug.Print x
```

## Breakpoint

Click left margin or press `F9`

## Step Through

`F8`

## Watch Variables

Debug → Add Watch

---

# 20. With Statement

```vba
With Range("A1")

    .Value = "Test"
    .Font.Bold = True
    .Interior.Color = vbYellow

End With
```

---

# 21. Events

## Workbook Open

```vba
Private Sub Workbook_Open()

    MsgBox "Workbook opened"

End Sub
```

## Worksheet Change

```vba
Private Sub Worksheet_Change(ByVal Target As Range)

    MsgBox "Cell changed"

End Sub
```

---

# 22. UserForms

## Show Form

```vba
UserForm1.Show
```

## Common Controls

| Control       | Purpose      |
| ------------- | ------------ |
| Label         | Display text |
| TextBox       | Input        |
| ComboBox      | Dropdown     |
| ListBox       | List         |
| CommandButton | Button       |
| CheckBox      | Toggle       |

---

# 23. Working with Files

## Open Workbook

```vba
Workbooks.Open "C:\Sales.xlsx"
```

## Save Workbook

```vba
ThisWorkbook.Save
```

## Save As

```vba
ThisWorkbook.SaveAs "C:\NewFile.xlsx"
```

## Close Workbook

```vba
Workbooks("Sales.xlsx").Close
```

---

# 24. File System Object

```vba
Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("C:\test.txt") Then
    MsgBox "Exists"
End If
```

---

# 25. Working with Tables

## Reference Table

```vba
Dim tbl As ListObject

Set tbl = ActiveSheet.ListObjects("SalesTable")
```

## Add Row

```vba
tbl.ListRows.Add
```

## Resize Table

```vba
tbl.Resize Range("A1:D100")
```

---

# 26. Filters & Sorting

## Apply Filter

```vba
Range("A1:D100").AutoFilter Field:=1, Criteria1:="East"
```

## Remove Filter

```vba
ActiveSheet.AutoFilterMode = False
```

## Sort

```vba
Range("A1:D100").Sort _
    Key1:=Range("A1"), _
    Order1:=xlAscending, _
    Header:=xlYes
```

---

# 27. Charts

## Create Chart

```vba
Dim ch As ChartObject

Set ch = ActiveSheet.ChartObjects.Add(100, 100, 400, 300)

ch.Chart.SetSourceData Source:=Range("A1:B10")
ch.Chart.ChartType = xlColumnClustered
```

---

# 28. Pivot Tables

## Create Pivot Cache

```vba
Set pc = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=Range("A1:D100"))
```

---

# 29. Named Ranges

## Create Name

```vba
Names.Add Name:="Sales", RefersTo:=Range("A1:A10")
```

## Use Name

```vba
Range("Sales").Select
```

---

# 30. Application Object

## Turn Off Screen Updating

```vba
Application.ScreenUpdating = False
```

## Disable Alerts

```vba
Application.DisplayAlerts = False
```

## Manual Calculation

```vba
Application.Calculation = xlCalculationManual
```

## Restore Settings

```vba
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
```

---

# 31. Best Practice Performance Pattern

```vba
Sub FastMacro()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Your code

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
```

---

# 32. Useful Built-in Functions

| Function    | Purpose       |
| ----------- | ------------- |
| UCase()     | Uppercase     |
| LCase()     | Lowercase     |
| Trim()      | Remove spaces |
| Left()      | Left chars    |
| Right()     | Right chars   |
| Mid()       | Extract text  |
| Len()       | String length |
| IsNumeric() | Check number  |
| IsDate()    | Check date    |
| Replace()   | Replace text  |
| Split()     | Split string  |

Example:

```vba
text = Replace("Hello World", "World", "Excel")
```

---

# 33. Date Functions

```vba
today = Date
nowTime = Now
yearVal = Year(Date)
monthVal = Month(Date)
dayVal = Day(Date)
```

---

# 34. Worksheet Functions in VBA

```vba
result = Application.WorksheetFunction.Sum(Range("A1:A10"))
```

Common ones:

| Function |
| -------- |
| Sum      |
| Average  |
| Count    |
| VLookup  |
| Match    |
| Index    |

---

# 35. Advanced Range Techniques

## Current Region

```vba
Range("A1").CurrentRegion
```

## Special Cells

```vba
Range("A1:A10").SpecialCells(xlCellTypeBlanks)
```

## Union

```vba
Union(Range("A1"), Range("C1"))
```

---

# 36. Common VBA Constants

| Constant  | Meaning        |
| --------- | -------------- |
| xlUp      | Up direction   |
| xlDown    | Down direction |
| xlToLeft  | Left           |
| xlToRight | Right          |
| xlYes     | Has headers    |
| xlNo      | No headers     |
| vbYes     | Yes button     |
| vbNo      | No button      |

---

# 37. Class Modules

## Basic Class

```vba
Private pName As String

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Name() As String
    Name = pName
End Property
```

---

# 38. API Calls (Advanced)

```vba
Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
```

---

# 39. Common Automation Tasks

## Loop Through Sheets

```vba
Dim ws As Worksheet

For Each ws In Worksheets
    Debug.Print ws.Name
Next ws
```

## Export PDF

```vba
ActiveSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:="C:\Report.pdf"
```

## Email via Outlook

```vba
Dim olApp As Object
Dim olMail As Object

Set olApp = CreateObject("Outlook.Application")
Set olMail = olApp.CreateItem(0)

With olMail
    .To = "test@email.com"
    .Subject = "Report"
    .Body = "Attached"
    .Display
End With
```

---

# 40. VBA Coding Standards

## Recommended

* Use `Option Explicit`
* Avoid `.Select`
* Use meaningful variable names
* Modularize code
* Add comments
* Handle errors properly

## Avoid

```vba
Range("A1").Select
Selection.Value = 5
```

Prefer:

```vba
Range("A1").Value = 5
```

---

# 41. Common VBA Interview Questions

## Difference Between Sub and Function

* `Sub` performs actions
* `Function` returns values

## What Is Variant?

A flexible data type that can hold any value.

## What Is Option Explicit?

```vba
Option Explicit
```

Forces variable declaration.

---

# 42. Essential Object Hierarchy

```text
Application
 └── Workbooks
      └── Worksheets
           └── Range
```

---

# 43. Frequently Used Snippets

## AutoFit Columns

```vba
Columns.AutoFit
```

## Hide Rows

```vba
Rows("5:10").Hidden = True
```

## Protect Sheet

```vba
ActiveSheet.Protect Password:="1234"
```

## Unprotect Sheet

```vba
ActiveSheet.Unprotect Password:="1234"
```

---

# 44. VBA Execution Flow

```text
Workbook Opens
   ↓
Workbook Events
   ↓
User Runs Macro
   ↓
Sub Calls Functions
   ↓
Manipulate Workbook
   ↓
Save / Export
```

---

# 45. Recommended Learning Path

1. VBA Editor basics
2. Variables & loops
3. Worksheets/ranges
4. Procedures/functions
5. Error handling
6. Arrays & collections
7. Events & UserForms
8. File automation
9. APIs & advanced automation
10. COM integration

---

# 46. High-Value Real-World Macros

## Import CSV Files

## Consolidate Reports

## Generate PDFs

## Send Automated Emails

## Clean Raw Data

## Build Dashboards

## Refresh Pivot Tables

## Create Invoice Systems

## Build Data Entry Forms

## Automate Monthly Reporting

---

# 47. VBA vs Modern Alternatives

| Tool           | Best For                 |
| -------------- | ------------------------ |
| VBA            | Excel desktop automation |
| Power Query    | Data transformation      |
| Office Scripts | Web automation           |
| Python         | Advanced analytics       |
| Power Automate | Workflow automation      |

---

# 48. Common Errors

| Error                     | Cause                        |
| ------------------------- | ---------------------------- |
| Subscript out of range    | Bad sheet/workbook reference |
| Object required           | Object not set               |
| Type mismatch             | Wrong data type              |
| Overflow                  | Number too large             |
| Application-defined error | Invalid Excel operation      |

---

# 49. Professional VBA Template

```vba
Option Explicit

Sub Main()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Code here

Cleanup:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Exit Sub

ErrorHandler:

    MsgBox Err.Description

    Resume Cleanup

End Sub
```

---

# 50. VBA Quick Reference

## Core Objects

| Object      | Description  |
| ----------- | ------------ |
| Application | Excel app    |
| Workbook    | Excel file   |
| Worksheet   | Sheet        |
| Range       | Cells        |
| Chart       | Chart object |
| ListObject  | Table        |

---

# 51. Useful Resources

* [Microsoft VBA Documentation](https://learn.microsoft.com/en-us/office/vba/api/overview/excel?utm_source=chatgpt.com)
* [Excel VBA Language Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference?utm_source=chatgpt.com)
* [Excel Object Model Reference](https://learn.microsoft.com/en-us/office/vba/api/excel.object-model?utm_source=chatgpt.com)
* [Rubberduck VBA Add-in](https://rubberduckvba.com?utm_source=chatgpt.com)

---

# 52. Most Important Habit

Always start modules with:

```vba
Option Explicit
```

It prevents countless bugs and makes VBA significantly more reliable.
