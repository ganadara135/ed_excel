Attribute VB_Name = "Module_main"
Option Explicit

Dim myChart As EmChartClass     '���������� �۵� ����
'Dim myChart2 As EmChartClass

Sub main()

Import_CSV_Copy
Sheets("EDChart").Activate
'Range("C1", Range("C1").End(xlDown)).Select
'Worksheets("EDChart").Range("C1").End(xlDown).Select

' ��Ʈ �׷��� ������, ������ ���ϱ�
MsgBoxCheck ("��Ʈ �׸���")
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
ActiveChart.SetSourceData Source:=Range("B2", Range("B2").End(xlDown))
'Dim myChart As EmChartClass
Set myChart = New EmChartClass
'activateChartEvent (myChart)


MsgBoxCheck ("��Į����: " & Range("B" & Rows.Count).End(xlUp).Row)

ResizeChart (Range("B" & Rows.Count).End(xlUp).Row)

MsgBoxCheck ("��Ʈ���� �������� �������� �ѹ��� Ŭ���ϼ���")
'ActiveChart.Activate



Range("C1").Value = "����"
Dim LR As Long, i As Long
LR = Range("B" & Rows.Count).End(xlUp).Row
For i = 3 To LR
    With Range("C" & i)
        .Value = Range("B" & i).Value - Range("B" & i - 1).Value
    End With
Next i

' ��Ʈ �׷��� ������ ����Ʈ ���ϱ�

Range("D1").Value = "������"

UserForm_Progress.CommandButton1.BackColor = &H8000000D
UserForm_Progress.Show vbModeless

DoEvents

End Sub

'Sub activateChartEvent()
' Set myChart = New EmChartClass
'End Sub

Sub ResizeChart(rowSize As Integer)
   With ActiveChart.Parent
      .Height = 200
      .Width = (1 * rowSize)
      .Top = 100
      .Left = 100
   End With
End Sub

 'Range("C19").Select
  '  Selection.Copy
   ' Range("F20").PasteSpecial xlPasteFormulas


Sub MsgBoxCheck(titleMsg As String)
   
    Dim MBNum As Integer
   
    MBNum = MsgBox(titleMsg & " �ܰ踦 �����ұ��?", vbYesNo, titleMsg)
   
    'MsgBox "�޽����ڽ� ���ϰ��� " & MBNum & "�Դϴ�"
    If (MBNum = 7) Then
        Exit Sub
    End If
   
End Sub
