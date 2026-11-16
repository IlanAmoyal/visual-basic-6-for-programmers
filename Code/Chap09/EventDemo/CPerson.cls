VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strName As String
Private m_nAge As Integer

Public Event Changed()
Public Event WillReset(ByRef bCancel As Boolean)

Private Sub Class_Initialize()
    m_strName = ""
    m_nAge = 0
End Sub

Public Property Let Name(strName As String)
    m_strName = strName
    RaiseEvent Changed
End Property

Public Property Get Name() As String
    Name = m_strName
End Property

Public Property Let Age(ByVal nAge As Integer)
    If nAge >= 0 Then
        m_nAge = nAge
        RaiseEvent Changed
    End If
End Property

Public Property Get Age() As Integer
    Age = m_nAge
End Property

Public Sub Reset()
    Dim bCnl As Boolean
    bCnl = False
    RaiseEvent WillReset(bCnl)
    If bCnl = False Then
        m_strName = ""
        m_nAge = 0
        RaiseEvent Changed
    End If
End Sub
