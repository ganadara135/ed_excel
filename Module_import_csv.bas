Attribute VB_Name = "Module1"
Option Explicit


Sub CSV_Import()
Dim ws As Worksheet
Dim FileName As Variant

Set ws = ActiveSheet 'set to current worksheet name

'Get the file name
FileName = Application.GetOpenFilename(FileFilter:="All Files (*.csv),*.csv", FilterIndex:=1, Title:="Select the CSV file", MultiSelect:=False)

If FileName = False Then Exit Sub

Application.ScreenUpdating = False

Application.StatusBar = "Reading the file... (" & FileName & ")"

'strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select text file...")

With ws.QueryTables.Add(Connection:="TEXT;" & FileName, Destination:=ws.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
End With

Application.StatusBar = "Reading the file... Done"
 
End Sub

