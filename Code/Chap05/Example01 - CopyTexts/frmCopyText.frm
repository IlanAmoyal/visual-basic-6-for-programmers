VERSION 5.00
Begin VB.Form frmCopyText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmCopyText"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIncludeText2 
      Caption         =   "chkIncludeText2"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox chkIncludeText1 
      Caption         =   "chkIncludeText1"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtInput2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "txtInput2"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.OptionButton optSeparateByNewLine 
      Caption         =   "optSeparateByNewLine"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.OptionButton optSeparateBySpace 
      Caption         =   "optSeparateBySpace"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdCopyTexts 
      Caption         =   "cmdCopyTexts"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox txtInput1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "txtInput1"
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtOutput 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmCopyText.frx":0000
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label lblInput2 
      AutoSize        =   -1  'True
      Caption         =   "lblInput2"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblInput1 
      AutoSize        =   -1  'True
      Caption         =   "lblInput1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmCopyText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmCopyText
'
' Project: CopyText
'
' Purpose: Demostrate use of check boxes and option buttons
'
' Copyright (C) 2001 by Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdCopyTexts_Click()
    ' Copy first input text only if chkInputText1 is checked
    If chkIncludeText1.Value = vbChecked Then
        txtOutput.Text = txtOutput.Text & txtInput1.Text
    End If
    
    ' Copy space or new-line only if txtOutput is not empty
    If txtOutput.Text <> "" Then
        ' Determines which option button is checked and copy new-line
        ' or space accordingly
        If optSeparateByNewLine.Value = True Then
            txtOutput.Text = txtOutput.Text & vbCrLf
        Else
            txtOutput.Text = txtOutput.Text & " "
        End If
    End If
    
    ' Copy second text only if chkInputText2 is checked
    If chkIncludeText2.Value = vbChecked Then
        txtOutput.Text = txtOutput.Text & txtInput2.Text
    End If
    
    ' Prepare the output text box to the next copy
    If txtOutput.Text <> "" Then
        txtOutput.Text = txtOutput.Text & vbCrLf
    End If
    
    ' Bring the focus back to input1
    txtInput1.SelStart = 0
    txtInput1.SelLength = Len(txtInput1.Text)
    txtInput1.SetFocus
End Sub

Private Sub Form_Load()
    ' Initialize the form caption
    frmCopyText.Caption = "Copy Text"
    
    ' Initialize the labels caption
    lblInput1.Caption = "Input &1"
    lblInput2.Caption = "Input &2"
    
    ' Clear the text boxes
    txtInput1.Text = ""
    txtInput2.Text = ""
    txtOutput.Text = ""
    
    ' Initialize the check boxes
    chkIncludeText1.Caption = "Copy to output"
    chkIncludeText1.Value = vbChecked
    chkIncludeText2.Caption = "Copy to output"
    chkIncludeText2.Value = vbChecked
    
    ' Initialize the option buttons
    optSeparateByNewLine.Caption = "Separate by a &new-line"
    optSeparateBySpace.Caption = "Separate by a &space"
    optSeparateBySpace = True
    
    ' Initialize the command button
    cmdCopyTexts.Caption = "&Copy texts to output"
End Sub
