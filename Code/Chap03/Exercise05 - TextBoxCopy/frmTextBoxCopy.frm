VERSION 5.00
Begin VB.Form frmTextBoxCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy between TextBoxes"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Terminate the program"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Clears the boxes"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtOutput 
      Height          =   1575
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Copy to output"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblInput 
      Caption         =   "Input"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTextBoxCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmTextBoxCopy.frm
'
' Project: TextBoxCopy
'
' Purpose: Demonstrate the use of concatenation, as well as SelStart and SelLength
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdClose_Click()
    Unload Me ' Terminate the program
End Sub

Private Sub cmdClear_Click()
    ' Clears the TextBoxes
    txtInput.Text = ""
    txtOutput.Text = ""
    txtInput.SetFocus
End Sub

Private Sub cmdCopy_Click()
    ' Add to the text in txtOutput, a new line with the text in txtInput
    txtOutput.Text = txtOutput.Text & txtInput.Text & vbCrLf
    ' Bring the focus back to txtInput
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    txtInput.SetFocus
End Sub

