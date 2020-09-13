Attribute VB_Name = "Module_main"
Option Explicit

Dim myChart As EmChartClass     '전역변수로 작동 안함
'Dim myChart2 As EmChartClass

Sub main()

Import_CSV_Copy
Sheets("EDChart").Activate
'Range("C1", Range("C1").End(xlDown)).Select
'Worksheets("EDChart").Range("C1").End(xlDown).Select

' 차트 그려서 시작점, 종료점 정하기
MsgBoxCheck ("차트 그리기")
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
ActiveChart.SetSourceData Source:=Range("B2", Range("B2").End(xlDown))
'Dim myChart As EmChartClass
Set myChart = New EmChartClass
'activateChartEvent (myChart)


MsgBoxCheck ("총칼럼수: " & Range("B" & Rows.Count).End(xlUp).Row)

ResizeChart (Range("B" & Rows.Count).End(xlUp).Row)

MsgBoxCheck ("차트에서 시작점과 종료점을 한번씩 클릭하세요")
'ActiveChart.Activate



Range("C1").Value = "차이"
Dim LR As Long, i As Long
LR = Range("B" & Rows.Count).End(xlUp).Row
For i = 3 To LR
    With Range("C" & i)
        .Value = Range("B" & i).Value - Range("B" & i - 1).Value
    End With
Next i

' 차트 그려서 변곡점 포인트 정하기

Range("D1").Value = "변곡점"

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
   
    MBNum = MsgBox(titleMsg & " 단계를 진행할까요?", vbYesNo, titleMsg)
   
    'MsgBox "메시지박스 리턴값은 " & MBNum & "입니다"
    If (MBNum = 7) Then
        Exit Sub
    End If
   
End Sub
