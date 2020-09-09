Attribute VB_Name = "Module1"
Option Explicit


Sub CSV_Import()
Dim ws As Worksheet, strFile As String

Set ws = ActiveSheet 'set to current worksheet name

strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please select text file...")

With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
End With
End Sub

