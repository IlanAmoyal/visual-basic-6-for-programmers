VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDynamicArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: CDynamicArray.cls
'
' Purpose:  Demonstrate the use of index inside property
'           The low index of the dynamic array is 0
'           The user of this class can enter value to any positive index,
'           if the given index is not allocated yet, it will be immedietly
'           after the call. If the dimension is negative, an error occur.
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private m_anValues() As Integer ' Intrinsic array that holds the values of the array

Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of items in the array."
    ' Description:
    ' This Property return the number of items in the array
        
    On Error GoTo Err_Handler
    Count = UBound(m_anValues) - LBound(m_anValues) + 1
    Exit Property
Err_Handler:
    Count = 0
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Free the memory occupied by the array."
    ' Description:
    ' Free the memory allocated to the intrinsic array
    Erase m_anValues
End Sub

Public Property Get Item(ByVal n As Integer) As Integer
Attribute Item.VB_Description = "Set/Returns the value of an item in the array."
Attribute Item.VB_UserMemId = 0
    ' Description:
    ' This property return the value of the n'th item in the array
    
    On Error GoTo Err_Handler
    Item = m_anValues(n)
    Exit Property
Err_Handler:
    Err.Raise Err.Number
End Property

Public Property Let Item(ByVal n As Integer, ByVal value As Integer)
    ' Description:
    ' This property enter the given to the n'th position in the array
    ' If the position is invalid, the array is re-dimension
    
    On Error GoTo Err_Handler
    m_anValues(n) = value
    Exit Property
Err_Handler:
    If n > 0 Then
        ReDim Preserve m_anValues(n)
        Resume
    Else
        Err.Raise Err.Number
    End If
End Property

