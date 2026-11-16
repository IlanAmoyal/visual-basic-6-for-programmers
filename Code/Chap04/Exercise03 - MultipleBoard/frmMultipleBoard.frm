VERSION 5.00
Begin VB.Form frmMultipleBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple Board"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print &Board"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblBoard 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "lblBoard"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   9255
   End
End
Attribute VB_Name = "frmMultipleBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmMultipleBoard
'
' Purpose: Demonstrate the use of loops and string concatenations
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim iCols As Integer, iRows As Integer
    Const cnMaxRows As Integer = 10, cnMaxCols As Integer = 10
    lblBoard.Caption = "          ----------------------------------------------------------------|" & vbCrLf
    lblBoard.Caption = lblBoard.Caption & "          |   001   002   003   004   005   006   007   008   009   010   |" & vbCrLf
    lblBoard.Caption = lblBoard.Caption & "----------|---------------------------------------------------------------|" & vbCrLf
    For iRows = 1 To cnMaxRows
        lblBoard.Caption = lblBoard.Caption & Format(iRows, "\|\ \ \ 000\ \ \ \|\ \ \ ")
        For iCols = 1 To cnMaxCols
            lblBoard.Caption = lblBoard.Caption & Format(iRows * iCols, "000\ \ \ ")
        Next
        lblBoard.Caption = lblBoard.Caption & "|" & vbCrLf
    Next
    lblBoard.Caption = lblBoard.Caption & "---------------------------------------------------------------------------" & vbCrLf
End Sub

Private Sub Form_Load()
    lblBoard.Caption = ""
End Sub
