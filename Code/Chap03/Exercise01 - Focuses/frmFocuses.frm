VERSION 5.00
Begin VB.Form frmFocuses 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Focuses"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForth 
      Caption         =   "Forth"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSecond 
      Caption         =   "Second"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdThird 
      Caption         =   "Third"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmFocuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmFocuses.frm
'
' Project: Focuses
'
' Purpose: Demonstrate the use of SetFocus method
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdFirst_Click()
    cmdsecond.SetFocus
End Sub

Private Sub cmdFirst_GotFocus()
    Cls
    Print cmdFirst.Name
    Caption = cmdFirst.Caption
End Sub

Private Sub cmdForth_Click()
    cmdFirst.SetFocus
End Sub

Private Sub cmdForth_GotFocus()
    Cls
    Print cmdForth.Name
    Caption = cmdForth.Caption
End Sub

Private Sub cmdsecond_Click()
    cmdThird.SetFocus
End Sub

Private Sub cmdsecond_GotFocus()
    Cls
    Print cmdsecond.Name
    Caption = cmdsecond.Caption
End Sub

Private Sub cmdThird_Click()
    cmdForth.SetFocus
End Sub

Private Sub cmdThird_GotFocus()
    Cls
    Print cmdThird.Name
    Caption = cmdThird.Caption
End Sub
