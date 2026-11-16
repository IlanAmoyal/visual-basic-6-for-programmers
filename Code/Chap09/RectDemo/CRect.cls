VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_dLength As Double
Private m_dWidth As Double

Private Sub Class_Initialize()
    m_dLength = 0#
    m_dWidth = 0#
End Sub

Public Property Let Length(ByVal dLength As Double)
    Debug.Print "Inside Property Let Length"
    If dLength >= 0# Then
        m_dLength = dLength
    End If
End Property

Public Property Get Length() As Double
    Debug.Print "Inside Property Get Length"
    Length = m_dLength
End Property

Public Property Let Width(dWidth As Double)
    Debug.Print "Inside Property Let Width"
    If dWidth >= 0# Then
        m_dWidth = dWidth
    End If
End Property

Public Property Get Width() As Double
    Debug.Print "Inside Property Get Width"
    Width = m_dWidth
End Property

Public Function Perimeter() As Double
    Perimeter = 2# * (m_dLength + m_dWidth)
End Function

Public Function Surface() As Double
    Surface = m_dLength * m_dWidth
End Function
