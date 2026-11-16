VERSION 5.00
Begin VB.Form frmCommandBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3
   ScaleMode       =   0  'User
   ScaleWidth      =   3
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBoard 
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmCommandBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmCommandBoard
'
' Purpose: Demonstrate how to create control-array in runtime
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdBoard_Click(Index As Integer)
    ' Remove the Command only if it isn't the first (index = 0) command
    If Index <> 0 Then Unload cmdBoard(Index)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    ' Initialize the caption
    Me.Caption = "Command Board"
    
    ' Reposition the first control
    With cmdBoard(0)
        .Caption = "&0"
        .Left = 0
        .Top = 0
        .Width = 1
        .Height = 1
        .Visible = True
    End With
    
    ' Load the controls from 2 to 9 (th first is already on the form)
    ' Then, reposition the loaded controls
    For i = 1 To 8
        Load cmdBoard(i)
        With cmdBoard(i)
            .Caption = "&" & Trim(Str(i))
            .Left = i Mod 3
            .Top = i \ 3
            .Width = 1
            .Height = 1
            .Visible = True
        End With
    Next i
End Sub
