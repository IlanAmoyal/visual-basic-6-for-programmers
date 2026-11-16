VERSION 5.00
Begin VB.Form frmFulfil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fulfil List Box"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "cmdRemove"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "cmdQuit"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ListBox lstItems 
      Height          =   2010
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "txtInput"
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "cmdInsert"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblTextCpation 
      AutoSize        =   -1  'True
      Caption         =   "Enter your text here:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "frmFulfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmFulfil
'
' Purpose: Demonstrate the use of ListBox
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdInsert_Click()
    If txtInput.Text <> "" Then lstItems.AddItem txtInput.Text
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    txtInput.SetFocus
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    If lstItems.ListIndex > -1 Then lstItems.RemoveItem lstItems.ListIndex
End Sub

Private Sub Form_Load()
    txtInput.Text = ""
    lstItems.Clear
    cmdInsert.Caption = "&Insert"
    cmdQuit.Caption = "&Quit"
    cmdRemove.Caption = "&Remove"
End Sub

Private Sub lstItems_DblClick()
    MsgBox "The item is: " & lstItems.List(lstItems.ListIndex)
End Sub
