Attribute VB_Name = "Module_main"
Option Explicit

Sub main()

CSV_Import_Copy
Sheets("EDChart").Activate
Range("C1", Range("C1").End(xlDown)).Select
'Worksheets("EDChart").Range("C1").End(xlDown).Select


End Sub


 'Range("C19").Select
  '  Selection.Copy
   ' Range("F20").PasteSpecial xlPasteFormulas
