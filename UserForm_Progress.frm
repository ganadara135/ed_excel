VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Progress 
   Caption         =   "UserForm1"
   ClientHeight    =   2352
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm_Progress.frx":0000
   StartUpPosition =   1  '������ ���
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
    Range("E1").Value = "�ø�"
    Range("F1").Value = "����"
    
    Dim vTarget As Integer
    vTarget = TextBox1.Value
    
    Dim LR As Long, i As Long, j As Long, k As Long
    LR = Range("B" & Rows.Count).End(xlUp).Row
    
    Range("E2").Value = "��"
    For i = 3 To LR
        Range("E" & i).Value = ""
        If Range("D" & i).Value = "������" Then
            If Range("E" & i - 1).Value = "��" Then
                Range("E" & i).Value = "��"
            Else
                Range("E" & i).Value = "��"
            End If
        Else
            Range("E" & i).Value = Range("E" & i - 1).Value
        End If
    Next i
    
    
    For j = 2 To LR
        Range("F" & j).Value = ""
        If Abs(Range("C" & j).Value) > vTarget Then
            Range("F" & j).Value = "��"
        Else
            Range("F" & j).Value = "��"
        End If
    Next j
    
    
    Range("G1").Value = "�ð�"
    Range("H1").Value = "����"
    Range("I1").Value = "����"
    Range("J1").Value = "����"
    
    For k = 2 To LR
        Range("G" & k).Value = ""
        If Range("E" & k).Value = "��" And Range("F" & k).Value = "��" Then
            Range("G" & k).Value = Range("B" & k).Value
        Else
            Range("G" & k).Value = ""
        End If
        
        Range("H" & k).Value = ""
        If Range("E" & k).Value = "��" And Range("F" & k).Value = "��" Then
            Range("H" & k).Value = Range("B" & k).Value
        Else
            Range("H" & k).Value = ""
        End If
        
        Range("I" & k).Value = ""
        If Range("E" & k).Value = "��" And Range("F" & k).Value = "��" Then
            Range("I" & k).Value = Range("B" & k).Value
        Else
            Range("I" & k).Value = ""
        End If
        
        Range("J" & k).Value = ""
        If Range("E" & k).Value = "��" And Range("F" & k).Value = "��" Then
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
    
    MsgBoxCheck ("������ ����ϱ�")
    Range("D1").Value = "������"
    
    Dim LR As Long, i As Long
    LR = Range("B" & Rows.Count).End(xlUp).Row
    For i = 3 To LR
        Range("D" & i).Value = ""
        If Abs(Range("C" & i).Value) > vTarget And Abs(Range("C" & i - 1).Value) < vTarget Then
            Range("D" & i).Value = "������"
        End If
    Next i
    
    UserForm_Progress.CommandButton4.BackColor = &H8000000D
End Sub

Private Sub CommandButton3_Click()
    '���� ��Ʈ �����
    'Worksheets("EDChart").ChartObjects(1).Chart.ChartArea.ClearFormats
    
    ' ��Ʈ �׷��� ������, ������ ���ϱ�
    MsgBoxCheck ("��Ʈ �׸���")
    Range("C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    ActiveChart.SetSourceData Source:=Range("C3", Range("C3").End(xlDown))
    Dim myChart2 As EmChartClass
    Set myChart2 = New EmChartClass
    
    MsgBoxCheck ("��Į����: " & Range("C" & Rows.Count).End(xlUp).Row)
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
