VERSION 5.00
Begin VB.Form frmConcatenate2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concatenate2"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConcatenate 
      Caption         =   "Concatenate"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   6255
   End
   Begin VB.TextBox txtOutput 
      Height          =   855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Concatenate Input1 and Input2"
      Top             =   1920
      Width           =   6255
   End
   Begin VB.TextBox txtInput2 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtInput1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      Caption         =   "Output"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label lblInput2 
      AutoSize        =   -1  'True
      Caption         =   "Input2"
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblInput1 
      AutoSize        =   -1  'True
      Caption         =   "Input1"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   450
   End
End
Attribute VB_Name = "frmConcatenate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmConcatenate2.frm
'
' Project: Concatenate2
'
' Purpose: Demonstrate the use of the & operator for concatenation and vbCrLf
'          constant
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdConcatenate_Click()
    txtOutput.Text = txtInput1.Text & vbCrLf & txtInput2.Text
    txtInput1.Text = ""
    txtInput2.Text = ""
    txtInput1.SetFocus
End Sub
