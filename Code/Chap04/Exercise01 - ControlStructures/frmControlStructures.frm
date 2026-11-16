VERSION 5.00
Begin VB.Form frmControlStructures 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Structures"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIIf 
      Caption         =   "IIf"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdIf 
      Caption         =   "If"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmControlStructures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmControlStructures
'
' Purpose: Demostrate the use of If...Then...Else, Select...Case, IIf
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdIf_Click()
    Dim nVal As Integer
    nVal = Val(InputBox("Enter value between 0 to 2"))
    If nVal = 0 Then
        MsgBox "The value is zero"
    ElseIf nVal = 1 Then
        MsgBox "The value is one"
    ElseIf nVal = 2 Then
        MsgBox "The value is two"
    Else
        MsgBox "Wrong value!"
    End If
End Sub

Private Sub cmdIIf_Click()
    Dim nVal As Integer
    nVal = Val(InputBox("Enter value between 0 to 2"))
    MsgBox (IIf(nVal = 0, "The value is zero", _
                IIf(nVal = 1, "The value is one", _
                    IIf(nVal = 2, "The value is two", _
                    "Wrong value!"))))
End Sub

Private Sub cmdSelect_Click()
    Dim nVal As Integer
    nVal = Val(InputBox("Enter value between 0 to 2"))
    Select Case nVal
    Case 0
        MsgBox "The value is zero"
    Case 1
        MsgBox "The value is one"
    Case 2
        MsgBox "The value is two"
    Case Else
        MsgBox "Wrong value!"
    End Select
End Sub
