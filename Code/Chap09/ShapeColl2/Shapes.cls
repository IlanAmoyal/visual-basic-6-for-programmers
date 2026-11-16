VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_coll As Collection

Private Sub Class_Initialize()
    Set m_coll = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_coll = Nothing
End Sub

Public Sub Add(Item As IShape, Optional Key As String = "")
Attribute Add.VB_Description = "Add shape to the collection"
    If Len(Key) > 0 Then
        m_coll.Add Item, Key
    Else
        m_coll.Add Item
    End If
End Sub

Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Remove a particular shape from the collection."
    m_coll.Remove Index
End Sub

Public Function Item(Index As Variant) As IShape
Attribute Item.VB_Description = "Return a particular shape from the collection."
Attribute Item.VB_UserMemId = 0
    Set Item = m_coll.Item(Index)
End Function

Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of shapes in the collection"
    Count = m_coll.Count
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_coll.[_NewEnum]
End Property
