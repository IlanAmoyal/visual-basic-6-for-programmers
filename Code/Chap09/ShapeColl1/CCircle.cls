VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PI As Double = 3.14

Private m_dRadius As Double

Public Property Get Radius() As Double
    Radius = m_dRadius
End Property

Public Property Let Radius(dRad As Double)
    If dRad >= 0 Then m_dRadius = dRad
End Property

Public Function Surface() As Double
    Surface = PI * m_dRadius * m_dRadius
End Function

Public Function Perimeter() As Double
    Perimeter = 2 * PI * m_dRadius
End Function

