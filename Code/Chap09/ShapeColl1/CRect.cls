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

Public Property Get Length() As Double
    Length = m_dLength
End Property

Public Property Let Length(dLen As Double)
    If dLen >= 0 Then m_dLength = dLen
End Property

Public Property Get Width() As Double
    Width = m_dWidth
End Property

Public Property Let Width(dWid As Double)
    If dWid >= 0 Then m_dWidth = dWid
End Property

Public Function Surface() As Double
    Surface = m_dLength * m_dWidth
End Function

Public Function Perimeter() As Double
    Perimeter = 2 * (m_dLength + m_dWidth)
End Function

