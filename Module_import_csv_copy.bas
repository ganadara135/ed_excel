Attribute VB_Name = "Module_import_csv"
Option Explicit


Sub CSV_Import_Copy()
Dim ws As Worksheet
Dim FileName As Variant
Dim sqlstring As String

Set ws = ActiveSheet 'set to current worksheet name
ws.Name = "Raw"

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



Dim FirstRow&, FirstCol&, SecondRow&, SecondCol&
Dim myUsedRange As Range

FirstRow = Cells.Find(What:="TIME", SearchDirection:=xlNext, SearchOrder:=xlByRows).Row

On Error Resume Next
FirstCol = Cells.Find(What:="TIME", SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column
If Err.Number <> 0 Then
    Err.Clear
    MsgBox _
    "There are horizontally merged cells on the sheet" & vbCrLf & _
    "that should be removed in order to locate the range.", 64, "Please unmerge all cells."
    Exit Sub
End If

Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(FirstRow, FirstCol).End(xlDown)).Copy
myUsedRange.Select

Worksheets.Add(Before:=ws).Name = "EDChart"
ActiveSheet.Paste

'Raw Sheet 로 다시 돌아감
Sheets("Raw").Activate
'Worksheets("Raw").Activate
SecondRow = Cells.Find(What:="W_SYS", SearchDirection:=xlNext, SearchOrder:=xlByRows).Row

On Error Resume Next
SecondCol = Cells.Find(What:="W_SYS", SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column
If Err.Number <> 0 Then
    
    MsgBox _
    "에러 발생 W_SYS : " & Err.Number & _
     vbCrLf
     
    Err.Clear
    
    Exit Sub
End If


Range(Cells(SecondRow, SecondCol), Cells(SecondRow, SecondCol).End(xlDown)).Copy _
Worksheets("EDChart").Range("B1")

End Sub



