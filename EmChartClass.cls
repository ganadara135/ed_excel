VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmChartClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents CEvents As Chart
Attribute CEvents.VB_VarHelpID = -1


Private Sub Class_Initialize()
Set CEvents = ActiveSheet.ChartObjects(1).Chart

    With CEvents
     '.Type = xlLineMarkers
     .HasTitle = True
     .ChartTitle.Text = "테스트 타이틀"
    End With

End Sub

Private Sub CEvents_Activate()
 MsgBox "The chart Events are working"
End Sub

Private Sub CEvents_BeforeDoubleClick(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean)
 MsgBox "BeforeDoubleClick : " & ElementID & " : " & Arg1
 'MsgBox "xlSeries : " & xlSeries
 If ElementID = xlSeries Then
    Cancel = True
    MsgBox "SeriesIndex : " & Arg1 & "  /  " & "PointIndex : " & Arg2
    If ActiveSheet.Range("K1").Value = -1 Or IsEmpty(ActiveSheet.Range("K1").Value) Then
        ActiveSheet.Range("K1").Value = Arg2
    ElseIf ActiveSheet.Range("K2").Value = -1 Or IsEmpty(ActiveSheet.Range("K2").Value) Then
        ActiveSheet.Range("K2").Value = Arg2
        UserForm_Progress.CommandButton2.BackColor = &H8000000D
    End If
 End If
End Sub

'Private Sub CEvents_Select(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long)
'    MsgBox "요소 : " & Arg2
'    ActiveSheet.Range("K1").Value = Arg2
'End Sub


'Private Sub CEvents_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
'    MsgBox Button & "-" & Shift & "-" & x & "-" & y
'End Sub


'Private Sub CEvents_SeriesChange(ByVal SeriesIndex As Long, ByVal PointIndex As Long)
'    MsgBox SeriesIndex & " - " & PointIndex
'End Sub

