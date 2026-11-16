VERSION 5.00
Begin VB.Form frmSlideColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slide Colors"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hsbBlue 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   5
      Top             =   3720
      Width           =   4935
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   3
      Top             =   3000
      Width           =   4935
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label lblBlue 
      AutoSize        =   -1  'True
      Caption         =   "&Blue"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label lblGreen 
      AutoSize        =   -1  'True
      Caption         =   "&Green"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label lblRed 
      AutoSize        =   -1  'True
      Caption         =   "&Red"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label lblRGBCaption 
      AutoSize        =   -1  'True
      Caption         =   "Mixed color"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblRGB 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmSlideColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmSlideColors.frm
'
' Project: SlideColors
'
' Purpose: Demonstrate the use of scroll bars
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub hsbBlue_Change()
lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub

Private Sub hsbGreen_Change()
lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub

Private Sub hsbRed_Change()
    lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub
