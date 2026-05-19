Yes. In VBA, the closest built-in equivalent to a Python `dict` is the `Scripting.Dictionary` object from the Microsoft Scripting Runtime.

Example:

```vba
Sub DictionaryExample()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Add items
    dict.Add "apple", 5
    dict.Add "banana", 10

    ' Access items
    MsgBox dict("apple")   ' 5

    ' Update item
    dict("apple") = 7

    ' Check existence
    If dict.Exists("banana") Then
        MsgBox "Banana exists"
    End If

    ' Loop through keys
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key & ": " & dict(key)
    Next key

End Sub
```

Python comparison:

| Python        | VBA                                               |
| ------------- | ------------------------------------------------- |
| `d = {}`      | `Set dict = CreateObject("Scripting.Dictionary")` |
| `d["a"] = 1`  | `dict("a") = 1`                                   |
| `"a" in d`    | `dict.Exists("a")`                                |
| `for k in d:` | `For Each key In dict.Keys`                       |
| `del d["a"]`  | `dict.Remove "a"`                                 |

A few useful methods/properties:

```vba
dict.Add key, value
dict.Exists(key)
dict.Remove key
dict.RemoveAll
dict.Keys
dict.Items
dict.Count
```

Two common ways to create one:

### Late binding (most portable)

No reference needed:

```vba
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

### Early binding (better IntelliSense)

In the VBA editor:

* Tools → References
* Check **Microsoft Scripting Runtime**

Then:

```vba
Dim dict As Scripting.Dictionary
Set dict = New Scripting.Dictionary
```

A nice feature: values can be anything — strings, numbers, arrays, collections, even other dictionaries.

Nested dictionary example:

```vba
Dim users As Object
Set users = CreateObject("Scripting.Dictionary")

Dim user As Object
Set user = CreateObject("Scripting.Dictionary")

user("age") = 30
user("city") = "Mexico City"

users("john") = user

MsgBox users("john")("city")
```

One caveat compared to Python:

* VBA dictionaries are not as fast or flexible as Python dicts.
* Key comparison behavior can matter (`TextCompare` vs `BinaryCompare`).

Example:

```vba
dict.CompareMode = vbTextCompare
```

This makes `"Apple"` and `"apple"` treated as the same key.
