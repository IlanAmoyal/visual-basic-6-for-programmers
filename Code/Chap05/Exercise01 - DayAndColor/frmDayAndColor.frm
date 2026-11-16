VERSION 5.00
Begin VB.Form frmDayAndColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Day and Color"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBlue 
      BackColor       =   &H00FF0000&
      Caption         =   "&Blue"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkGreen 
      BackColor       =   &H0000FF00&
      Caption         =   "&Green"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkRed 
      BackColor       =   &H000000FF&
      Caption         =   "&Red"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame fraDay 
      Caption         =   "Day"
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton optSaturday 
         Caption         =   "S&aturday"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton optFriday 
         Caption         =   "&Friday"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optThursday 
         Caption         =   "T&hursday"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optWednsday 
         Caption         =   "&Wednsday"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optTuesday 
         Caption         =   "&Tuesday"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optMonday 
         Caption         =   "&Monday"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optSunday 
         Caption         =   "&Sunday"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   465
      Left            =   2265
      TabIndex        =   11
      Top             =   2520
      Width           =   1245
   End
End
Attribute VB_Name = "frmDayAndColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmDayAndColor.frm
'
' Purpose: Demonstrate the use of OptionButton, CheckBox and Frame
'
' Copyright (C) 2001 by Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub chkBlue_Click()
    ' Using bit-mask enables or disables the blue ingredient
    If chkBlue.Value = vbChecked Then
        BackColor = BackColor Or RGB(0, 0, 255)
    Else
        BackColor = BackColor And Not RGB(0, 0, 255)
    End If
End Sub

Private Sub chkGreen_Click()
    ' Using bit-mask enables or disables the green ingredient
    If chkGreen.Value = vbChecked Then
        BackColor = BackColor Or RGB(0, 255, 0)
    Else
        BackColor = BackColor And Not RGB(0, 255, 0)
    End If
End Sub

Private Sub chkRed_Click()
    ' Using bit-mask enables or disables the red ingredient
    If chkRed.Value = vbChecked Then
        BackColor = BackColor Or RGB(255, 0, 0)
    Else
        BackColor = BackColor And Not RGB(255, 0, 0)
    End If
End Sub

Private Sub Form_Load()
    ' When loading the program, reset all the ingredients of the color
    BackColor = RGB(0, 0, 0)
End Sub

Private Sub optFriday_Click()
    lblDay.Caption = "Friday"
End Sub

Private Sub optSaturday_Click()
    lblDay.Caption = "Saturday"
End Sub

Private Sub optThursday_Click()
    lblDay.Caption = "Thursday"
End Sub

Private Sub optWednsday_Click()
    lblDay.Caption = "Wednsday"
End Sub

Private Sub optMonday_Click()
    lblDay.Caption = "Monday"
End Sub

Private Sub optSunday_Click()
    lblDay.Caption = "Sunday"
End Sub

Private Sub optTuesday_Click()
    lblDay.Caption = "Thursday"
End Sub
