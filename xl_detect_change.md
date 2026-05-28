Use the worksheet’s `Change` event. In VBA, Excel can automatically run code whenever a cell is edited.

Example: trigger a macro when cell `A1` changes.

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    ' Check if the changed cell is A1
    If Not Intersect(Target, Me.Range("A1")) Is Nothing Then
        
        ' Prevent recursive triggering
        Application.EnableEvents = False
        
        ' Run your macro
        Call MyMacro
        
        Application.EnableEvents = True
    End If

End Sub
```

Then put your macro in a standard module:

```vb
Sub MyMacro()

    MsgBox "A1 changed!"

End Sub
```

How to add it:

1. Press `ALT + F11`
2. In the VBA editor, double-click the worksheet you want to monitor (for example `Sheet1`)
3. Paste the `Worksheet_Change` code there
4. Insert → Module
5. Paste the `MyMacro` procedure into the module

Notes:

* `Worksheet_Change` fires only when a user edits a cell or VBA changes it.
* Formula recalculations do **not** trigger this event. For formulas changing due to recalculation, use `Worksheet_Calculate` instead.
* `Application.EnableEvents = False` prevents infinite loops if your macro modifies cells.

Example for monitoring multiple cells:

```vb
If Not Intersect(Target, Me.Range("A1:A10")) Is Nothing Then
    Call MyMacro
End If
```

Example for a whole column:

```vb
If Not Intersect(Target, Me.Columns("B")) Is Nothing Then
    Call MyMacro
End If
```




You can monitor multiple cells by using `Intersect` with a larger range, multiple ranges, or an array of addresses.

## Option 1 — Monitor a continuous range

Example: monitor `A1:A10`

```vb id="80c6i7"
Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Target, Me.Range("A1:A10")) Is Nothing Then
        
        Application.EnableEvents = False
        
        Call MyMacro
        
        Application.EnableEvents = True
    End If

End Sub
```

---

## Option 2 — Monitor non-adjacent cells

Example: monitor `A1`, `C5`, and `F10`

```vb id="42f0mj"
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim WatchRange As Range
    
    Set WatchRange = Union( _
        Me.Range("A1"), _
        Me.Range("C5"), _
        Me.Range("F10") _
    )

    If Not Intersect(Target, WatchRange) Is Nothing Then
        
        Application.EnableEvents = False
        
        Call MyMacro
        
        Application.EnableEvents = True
    End If

End Sub
```

---

## Option 3 — Detect exactly which cell changed

```vb id="7obn6d"
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim c As Range

    For Each c In Target

        Select Case c.Address

            Case "$A$1"
                MsgBox "A1 changed"

            Case "$C$5"
                MsgBox "C5 changed"

            Case "$F$10"
                MsgBox "F10 changed"

        End Select

    Next c

End Sub
```

This is useful when different cells should trigger different actions.

---

## Option 4 — Monitor an entire table/column

```vb id="p7k0l3"
If Not Intersect(Target, Me.Columns("B")) Is Nothing Then
    Call MyMacro
End If
```

Or:

```vb id="t4nfg4"
If Not Intersect(Target, Me.Range("B:B,D:D")) Is Nothing Then
    Call MyMacro
End If
```

---

## Important tip for pasted values

If the user pastes multiple cells at once, `Target` may contain many cells. This pattern handles that safely:

```vb id="7n1o7t"
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim c As Range

    If Intersect(Target, Me.Range("A1:A10")) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    For Each c In Intersect(Target, Me.Range("A1:A10"))
        Debug.Print c.Address & " changed to " & c.Value
    Next c

    Application.EnableEvents = True

End Sub
```


