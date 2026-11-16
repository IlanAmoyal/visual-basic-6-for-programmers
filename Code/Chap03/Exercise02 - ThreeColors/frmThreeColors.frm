VERSION 5.00
Begin VB.Form frmThreeColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Three Colors"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBlue 
      Caption         =   "&Blue"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreen 
      Caption         =   "&Green"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "&Red"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmThreeColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmThreeColors.frm
'
' Project: ThreeColors
'
' Purpose: Demonstrate the use of RGB function,  and the BackColor property
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdBlue_Click()
    BackColor = RGB(0, 0, 255)
    Caption = "ThreeColors - Blue Now..."
End Sub

Private Sub cmdGreen_Click()
    BackColor = RGB(0, 255, 0)
    Caption = "ThreeColors - Green Now..."
End Sub

Private Sub cmdRed_Click()
    BackColor = RGB(255, 0, 0)
    Caption = "ThreeColors - Red Now..."
End Sub

