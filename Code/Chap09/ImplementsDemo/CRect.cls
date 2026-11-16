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

Implements IShape

Public m_dLength As Double
Public m_dWidth As Double

Private Function IShape_Perimeter() As Double
    IShape_Perimeter = 2 * (m_dLength + m_dWidth)
End Function

Private Function IShape_Surface() As Double
    IShape_Surface = m_dLength * m_dWidth
End Function
