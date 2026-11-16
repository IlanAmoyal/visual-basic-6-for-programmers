VERSION 5.00
Begin VB.Form frmSlideColors2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slide Colors 2"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtBlue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtRed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   375
      LargeChange     =   10
      Left            =   240
      Max             =   255
      TabIndex        =   5
      Top             =   3720
      Width           =   4935
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   375
      LargeChange     =   10
      Left            =   240
      Max             =   255
      TabIndex        =   3
      Top             =   3000
      Width           =   4935
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   375
      LargeChange     =   10
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
      Width           =   5775
   End
End
Attribute VB_Name = "frmSlideColors2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmSlideColors2.frm
'
' Project: SlideColors2
'
' Purpose: Demonstrate the use of scroll bars
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
    lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    txtRed = Str(hsbRed.Value)
    txtGreen.Text = Str(hsbGreen.Value)
    txtBlue.Text = Str(hsbBlue.Value)
    txtRed.BackColor = RGB(hsbRed.Value, 0, 0)
    txtGreen.BackColor = RGB(0, hsbGreen.Value, 0)
    txtBlue.BackColor = RGB(0, 0, hsbBlue.Value)
End Sub

Private Sub hsbBlue_Change()
    lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    txtBlue.Text = Str(hsbBlue.Value)
    txtBlue.BackColor = RGB(0, 0, hsbBlue.Value)
End Sub

Private Sub hsbGreen_Change()
    lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    txtGreen.Text = Str(hsbGreen.Value)
    txtGreen.BackColor = RGB(0, hsbGreen.Value, 0)
End Sub

Private Sub hsbRed_Change()
    lblRGB.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
    txtRed.Text = Str(hsbRed.Value)
    txtRed.BackColor = RGB(hsbRed.Value, 0, 0)
End Sub
