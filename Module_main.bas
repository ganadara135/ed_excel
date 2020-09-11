Attribute VB_Name = "Module_main"
Option Explicit

Sub main()

Import_CSV_Copy
Sheets("EDChart").Activate
'Range("C1", Range("C1").End(xlDown)).Select
'Worksheets("EDChart").Range("C1").End(xlDown).Select

Range("C1").Value = "Â÷ÀÌ"
Dim LR As Long, i As Long
LR = Range("B" & Rows.Count).End(xlUp).Row
For i = 3 To LR
    With Range("C" & i)
        .Value = Range("B" & i).Value - Range("B" & i - 1).Value
    End With
Next i


End Sub


 'Range("C19").Select
  '  Selection.Copy
   ' Range("F20").PasteSpecial xlPasteFormulas
