VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Progress 
   Caption         =   "UserForm1"
   ClientHeight    =   2352
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm_Progress.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub CommandButton5_Click()
    Range("E1").Value = "플마"
    Range("F1").Value = "고저"
    
    Dim vTarget As Integer
    vTarget = TextBox1.Value
    
    Dim LR As Long, i As Long, j As Long, k As Long
    LR = Range("B" & Rows.Count).End(xlUp).Row
    
    Range("E2").Value = "마"
    For i = 3 To LR
        Range("E" & i).Value = ""
        If Range("D" & i).Value = "변곡점" Then
            If Range("E" & i - 1).Value = "마" Then
                Range("E" & i).Value = "플"
            Else
                Range("E" & i).Value = "마"
            End If
        Else
            Range("E" & i).Value = Range("E" & i - 1).Value
        End If
    Next i
    
    
    For j = 2 To LR
        Range("F" & j).Value = ""
        If Abs(Range("C" & j).Value) > vTarget Then
            Range("F" & j).Value = "고"
        Else
            Range("F" & j).Value = "저"
        End If
    Next j
    
    
    Range("G1").Value = "플고"
    Range("H1").Value = "플저"
    Range("I1").Value = "마고"
    Range("J1").Value = "마저"
    
    For k = 2 To LR
        Range("G" & k).Value = ""
        If Range("E" & k).Value = "플" And Range("F" & k).Value = "고" Then
            Range("G" & k).Value = Range("B" & k).Value
        Else
            Range("G" & k).Value = ""
        End If
        
        Range("H" & k).Value = ""
        If Range("E" & k).Value = "플" And Range("F" & k).Value = "저" Then
            Range("H" & k).Value = Range("B" & k).Value
        Else
            Range("H" & k).Value = ""
        End If
        
        Range("I" & k).Value = ""
        If Range("E" & k).Value = "마" And Range("F" & k).Value = "고" Then
            Range("I" & k).Value = Range("B" & k).Value
        Else
            Range("I" & k).Value = ""
        End If
        
        Range("J" & k).Value = ""
        If Range("E" & k).Value = "마" And Range("F" & k).Value = "저" Then
            Range("J" & k).Value = Range("B" & k).Value
        Else
            Range("J" & k).Value = ""
        End If
    Next k
    
    
    UserForm_Progress.CommandButton4.BackColor = &H8000000D
End Sub

Private Sub CommandButton4_Click()
    Dim vTarget As Integer
    vTarget = TextBox1.Value
    
    MsgBoxCheck ("변곡점 계산하기")
    Range("D1").Value = "변곡점"
    
    Dim LR As Long, i As Long
    LR = Range("B" & Rows.Count).End(xlUp).Row
    For i = 3 To LR
        Range("D" & i).Value = ""
        If Abs(Range("C" & i).Value) > vTarget And Abs(Range("C" & i - 1).Value) < vTarget Then
            Range("D" & i).Value = "변곡점"
        End If
    Next i
    
    UserForm_Progress.CommandButton4.BackColor = &H8000000D
End Sub

Private Sub CommandButton3_Click()
    '기존 차트 지우기
    'Worksheets("EDChart").ChartObjects(1).Chart.ChartArea.ClearFormats
    
    ' 차트 그려서 시작점, 종료점 정하기
    MsgBoxCheck ("차트 그리기")
    Range("C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    ActiveChart.SetSourceData Source:=Range("C3", Range("C3").End(xlDown))
    Dim myChart2 As EmChartClass
    Set myChart2 = New EmChartClass
    
    MsgBoxCheck ("총칼럼수: " & Range("C" & Rows.Count).End(xlUp).Row)
    'MsgBox Range("C3", Range("C3").End(xlDown)).Rows
    
    ResizeChart (Range("C" & Rows.Count).End(xlUp).Row)
    
    UserForm_Progress.CommandButton3.BackColor = &H8000000D
End Sub


Sub ResizeChart(rowSize As Integer)
   With ActiveChart.Parent
      .Height = 200
      .Width = (1 * rowSize)
      .Top = 100
      .Left = 100
   End With
End Sub





Private Sub TextBox1_Change()

End Sub
