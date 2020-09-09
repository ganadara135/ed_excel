Attribute VB_Name = "Module_import_csv"
Option Explicit


Sub CSV_Import_Copy()
Dim ws As Worksheet
Dim FileName As Variant
Dim sqlstring As String

Set ws = ActiveSheet 'set to current worksheet name

'Get the file name
FileName = Application.GetOpenFilename(FileFilter:="All Files (*.csv),*.csv", FilterIndex:=1, Title:="Select the CSV file", MultiSelect:=False)

If FileName = False Then Exit Sub

Application.ScreenUpdating = False

Application.StatusBar = "Reading the file... (" & FileName & ")"

'strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select text file...")
sqlstring = "select DATE from 96Sales where profit < 5"

With ws.QueryTables.Add(Connection:="TEXT;" & FileName, Destination:=ws.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
End With

Application.StatusBar = "Reading the file... Done"
Application.ScreenUpdating = True


'Dim myUsedRange As Range
'Dim LastRow As Long, LastColumn As Long
'LastRow = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'LastColumn = Cells.Find(What:="W_SYS", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
'Range("A1").Resize(LastRow, LastColumn).Select
  'MsgBox "The data range address is " & Selection.Address(0, 0) & ".", 64, "Data-containing range address:"
'MsgBox "LastRow : " & LastRow & "LastColumn : " & LastColumn



Dim FirstRow&, FirstCol&, LastRow&, LastCol&
Dim myUsedRange As Range
FirstRow = Cells.Find(What:="W_SYS", SearchDirection:=xlNext, SearchOrder:=xlByRows).Row

On Error Resume Next
FirstCol = Cells.Find(What:="W_SYS", SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column
If Err.Number <> 0 Then
    Err.Clear
    MsgBox _
    "There are horizontally merged cells on the sheet" & vbCrLf & _
    "that should be removed in order to locate the range.", 64, "Please unmerge all cells."
    Exit Sub
End If

'LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
'LastCol = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
'Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(.Row, .CurrentRegion.Column))
'myUsedRange.Select
Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(FirstRow, FirstCol).End(xlDown)).Copy
myUsedRange.Select

Worksheets.Add(Before:=ws).Name = "EDChart"
ActiveSheet.Paste

End Sub



