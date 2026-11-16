VERSION 5.00
Begin VB.Form frmRectDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rectangle Demo"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetLength 
      Caption         =   "Get length"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetNew 
      Caption         =   "Set New"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdSetLength 
      Caption         =   "Set length"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetWidth 
      Caption         =   "Get width"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetWidth 
      Caption         =   "Set width"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmRectDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rc As CRect

Private Sub cmdGetLength_Click()
    MsgBox "Length = " & Str(rc.Length)
End Sub

Private Sub cmdGetWidth_Click()
    MsgBox "Width = " & Str(rc.Width)
End Sub

Private Sub cmdSetLength_Click()
    rc.Length = Val(InputBox("Enter the length:"))
End Sub

Private Sub cmdSetNew_Click()
    Set rc = New CRect
End Sub

Private Sub cmdSetWidth_Click()
    rc.Width = Val(InputBox("Enter the width:"))
End Sub
