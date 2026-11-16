VERSION 5.00
Begin VB.Form frmTestDArray 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmTestClass"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "cmdClear"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "cmdCount"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLet 
      Caption         =   "cmdLet"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "cmdGet"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "txtValue"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtIndex 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "txtIndex"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Index"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "frmTestDArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: cls...
'
' Purpose:
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Declare an object of CDynamicArray to test in this program
Private m_DArray As DynamicArray.CDynamicArray

Private Sub cmdClear_Click()
    ' Description:
    ' Clear all the items from the dynamic array using the clear method
    
    m_DArray.Clear
End Sub

Private Sub cmdCount_Click()
    ' Description:
    ' Display the size of the dynamic array using the Count method
    
    MsgBox "The array contains " & Str(m_DArray.Count) & " items"
End Sub

Private Sub cmdGet_Click()
    ' Description:
    ' Use the Item property to get the item whose index placed within txtIndex
    
    txtValue.Text = Str(m_DArray(Val(txtIndex.Text)))
End Sub

Private Sub cmdLet_Click()
    ' Description:
    ' Insert a new item to the array using the property Item
    
    m_DArray(Val(txtIndex.Text)) = Val(txtValue.Text)
End Sub

Private Sub Form_Load()
    ' Initializations:
    Me.Caption = "Test Dynamic Array"
    txtIndex.Text = ""
    txtValue.Text = ""
    cmdGet.Caption = "&Get"
    cmdLet.Caption = "&Let"
    cmdCount.Caption = "&Count"
    cmdClear.Caption = "Clea&r"
    
    ' Create the dynamic-array class
    Set m_DArray = New CDynamicArray
End Sub
