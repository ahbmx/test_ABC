In VBA, you can format a WWPN like:

```text
10000000C9AA1122
```

into:

```text
10:00:00:00:C9:AA:11:22
```

using a simple function.

```vba
Function FormatWWPN(ByVal wwpn As String) As String

    Dim i As Integer
    Dim result As String

    ' Remove existing colons/spaces
    wwpn = Replace(wwpn, ":", "")
    wwpn = Replace(wwpn, " ", "")

    ' Add colon every 2 characters
    For i = 1 To Len(wwpn) Step 2
        result = result & Mid(wwpn, i, 2) & ":"
    Next i

    ' Remove last colon
    If Right(result, 1) = ":" Then
        result = Left(result, Len(result) - 1)
    End If

    FormatWWPN = UCase(result)

End Function
```

Usage:

```vba
Sub Test()

    Dim rawWWPN As String

    rawWWPN = "10000000c9aa1122"

    MsgBox FormatWWPN(rawWWPN)

End Sub
```

Result:

```text
10:00:00:00:C9:AA:11:22
```

You can also use it directly in Excel cells:

```excel
=FormatWWPN(A1)
```

if the VBA function is in a module.

A shorter version:

```vba
Function FormatWWPN(s As String) As String

    Dim i As Long

    s = Replace(s, ":", "")

    For i = 2 To Len(s) - 1 Step 2
        s = Left(s, i - 1) & ":" & Mid(s, i)
        i = i + 1
    Next i

    FormatWWPN = UCase(s)

End Function
```

The first version is cleaner and safer for storage admin work.
