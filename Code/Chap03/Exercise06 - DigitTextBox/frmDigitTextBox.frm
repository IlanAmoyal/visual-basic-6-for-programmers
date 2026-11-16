VERSION 5.00
Begin VB.Form frmDigitTextBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Only digit enable TextBox"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Clears the text"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Terminate the program"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtDigits 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmDigitTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmDigitTextBox.frm
'
' Project: DigitTextBox
'
' Purpose: Demonstrate how to control the input using the KeyPress event
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdClear_Click()
    txtDigits.Text = "" ' Clears the TextBoxes
    txtDigits.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me ' Terminate the program
End Sub

Private Sub txtDigits_KeyPress(KeyAscii As Integer)
    ' Check whether the character is not between "0" to "9" and also
    ' not a backspace or newline(return) - if so, cancel the printing
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
        KeyAscii <> vbKeyBack And _
        KeyAscii <> vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
