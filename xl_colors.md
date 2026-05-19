Here's the complete **Excel ColorIndex Reference Guide** (1-56):

## 🎨 Basic Colors

| Index | Color | Index | Color |
|-------|-------|-------|-------|
| 1 | ⬛ Black | 2 | ◻️ White |
| 3 | 🟥 Red | 4 | 🟩 Bright Green |
| 5 | 🟦 Blue | 6 | 🟨 Yellow |
| 7 | 🩷 Pink | 8 | 💙 Turquoise |
| 9 | 🟥 Dark Red | 10 | 🟢 Green |
| 11 | 🔵 Dark Blue | 12 | 🟡 Dark Yellow |
| 13 | 🟣 Violet | 14 | 🟦 Teal |
| 15 | ⬜ Gray | 16 | ⬛ Dark Gray |
| 17 | 🟤 Light Brown | 18 | 🟤 Brown |
| 19 | 🟤 Dark Brown | 20 | 🫒 Olive Green |

## 🌈 Extended Colors

| Index | Color | Index | Color |
|-------|-------|-------|-------|
| 21 | 🔷 Bright Blue | 22 | 🔵 Dark Blue |
| 23 | 💚 Light Turquoise | 24 | 🟣 Indigo |
| 25 | 💜 Periwinkle | 26 | 🟣 Plum |
| 27 | 🦷 Ivory | 28 | 💛 Light Yellow |
| 29 | 🫒 Olive | 30 | 🫒 Dark Olive |
| 31 | 🩵 Blue-Gray | 32 | 🧡 Orange |
| 33 | 🟠 Light Orange | 34 | 🟢 Lime |
| 35 | ✨ Gold | 36 | 🍑 Peach |
| 37 | 💚 Light Green | 38 | 🌊 Sea Green |
| 39 | 💙 Aqua | 40 | ☁️ Sky Blue |
| 41 | 🤎 Tan | 42 | 💛 Light Gold |

## 🌟 Vivid & Pastel Colors

| Index | Color | Index | Color |
|-------|-------|-------|-------|
| 43 | 🟣 Lavender | 44 | 💙 Light Blue |
| 45 | 🟢 Dark Teal | 46 | 🧡 Orange |
| 47 | 🩷 Pink | 48 | 🟣 Plum |
| 49 | 🌸 Rose | 50 | 🧸 Beige |
| 51 | 🌿 Mint | 52 | 🪸 Coral |
| 53 | 🐟 Salmon | 54 | 🏜️ Tan |
| 55 | 🌾 Wheat | 56 | 🧈 Ivory |

## 📋 Quick Copy - 16 Recommended Colors for Your Groups

If you want 16 distinct, pleasant colors for your groups:

```vbnet
Group 1:  ColorIndex = 3   ' Red
Group 2:  ColorIndex = 4   ' Bright Green
Group 3:  ColorIndex = 5   ' Blue
Group 4:  ColorIndex = 6   ' Yellow
Group 5:  ColorIndex = 7   ' Pink
Group 6:  ColorIndex = 8   ' Turquoise
Group 7:  ColorIndex = 9   ' Dark Red
Group 8:  ColorIndex = 10  ' Green
Group 9:  ColorIndex = 11  ' Dark Blue
Group 10: ColorIndex = 13  ' Violet
Group 11: ColorIndex = 14  ' Teal
Group 12: ColorIndex = 17  ' Light Brown
Group 13: ColorIndex = 32  ' Orange
Group 14: ColorIndex = 34  ' Lime
Group 15: ColorIndex = 46  ' Bright Orange
Group 16: ColorIndex = 50  ' Beige
```

## 💡 Pro Tip

To see all 56 colors visually in Excel:
1. Open a blank worksheet
2. Run this quick macro (or just type the numbers in cells A1-A56 and manually fill colors):

```vbnet
Sub ShowAllColors()
    Dim i As Integer
    For i = 1 To 56
        Cells(i, 1).Interior.ColorIndex = i
        Cells(i, 2).Value = i
    Next i
End Sub
```

This will fill column A with each color and show the index number in column B.

---

Which 16 colors would you like me to set for your groups? Just tell me the index numbers!
