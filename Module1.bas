Attribute VB_Name = "Module1"
Option Explicit


'
' ��ũ��1 ��ũ��
'

'
'    Range("B2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
'    ActiveChart.SetSourceData Source:=Range("EDChart!$B$2:$B$2510")

Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _
        ByVal PointIndex As Long)
'    Set p = ActiveChart.SeriesCollection(SeriesIndex). _
 '       Points(PointIndex)
  '  p.Border.ColorIndex = 3
  msg "call Char_SeriesChange()"
End Sub

