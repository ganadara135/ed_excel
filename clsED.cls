VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'A class module has two events
'Initialize  occurs when a new object of the class is created.
'Terminate  occurrs when the class object is deleted.

'Let is used to assign a value to a basic variable type.
'Set is used to assign an object to an object variable.


' clsED class module
Private m_Sys As Long
Private m_Time As Time
' Properties
Public Property Get getSys() As Long
    getSys = m_Sys
End Property
Public Property Let letSys(ByVal vSys As Long)
    m_Sys = vSys
End Property
Public Property Get getTime() As Time
    getTime = m_Time
End Property
Public Property Let letTime(ByVal vTime As Time)
    m_Time = vTime
End Property
