VERSION 5.00
Begin VB.Form frmClassDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Class Demo"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReleaseD3 
      Caption         =   "Release d3"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetValOfD3 
      Caption         =   "Set val of d3"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetValOfD3 
      Caption         =   "Get val of d3"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetValOfD2 
      Caption         =   "Set val of d2"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetValOfD1 
      Caption         =   "Get val of d1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdReleaseD1 
      Caption         =   "Release d1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetValOfD1 
      Caption         =   "Set val of d1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdReleaseD2 
      Caption         =   "Release d2"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetValOfD2 
      Caption         =   "Get val of d2"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdInitializeD3 
      Caption         =   "Initialize d3"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdInitializeD2 
      Caption         =   "Initialize d2"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmClassDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim d1 As New CDemo
Dim d2 As CDemo
Dim d3 As CDemo

Private Sub cmdGetValOfD1_Click()
    MsgBox "d1.val = " & Str(d1.val)
End Sub

Private Sub cmdGetValOfD2_Click()
    MsgBox "d2.val = " & Str(d2.val)
End Sub

Private Sub cmdGetValOfD3_Click()
    MsgBox "d3.val = " & Str(d3.val)
End Sub

Private Sub cmdInitializeD2_Click()
    Set d2 = New CDemo
End Sub

Private Sub cmdInitializeD3_Click()
    Set d3 = d2
End Sub

Private Sub cmdReleaseD1_Click()
    Set d1 = Nothing
End Sub

Private Sub cmdReleaseD2_Click()
    Set d2 = Nothing
End Sub

Private Sub cmdReleaseD3_Click()
    Set d3 = Nothing
End Sub

Private Sub cmdSetValOfD1_Click()
    d1.val = 100
End Sub

Private Sub cmdSetValOfD2_Click()
    d2.val = 200
End Sub

Private Sub cmdSetValOfD3_Click()
    d3.val = 300
End Sub

