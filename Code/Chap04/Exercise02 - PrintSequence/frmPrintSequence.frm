VERSION 5.00
Begin VB.Form frmPrintSequence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Squence"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCloseProgram 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtNum 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Text            =   "txtNum"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblSquence 
      Caption         =   "lblSquence"
      Height          =   4935
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
End
Attribute VB_Name = "frmPrintSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmPrintSequence
'
' Purpose: Demonstate the use of loops (in this case the For...Next)
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdCloseProgram_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim strSequence As String
    Dim nVal As Integer, nCounter As Integer
    nVal = Val(txtNum.Text)
    strSequence = "The sequence is:" & vbCrLf
    For nCounter = 0 To nVal
        strSequence = strSequence & str(nCounter) & " "
    Next
    lblSquence.Caption = strSequence
    txtNum.SelStart = 0
    txtNum.SelLength = Len(txtNum.Text)
    txtNum.SetFocus
End Sub

Private Sub Form_Load()
    lblSquence = "The squence is: " & vbCrLf
    txtNum.Text = ""
End Sub
