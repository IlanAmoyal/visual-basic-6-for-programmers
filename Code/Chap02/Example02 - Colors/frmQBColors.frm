VERSION 5.00
Begin VB.Form frmQBColors 
   Caption         =   "QBColors"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   6600
   Begin VB.Menu mnuColors 
      Caption         =   "&Colors"
      Begin VB.Menu mnuColorsRed 
         Caption         =   "&Red"
      End
      Begin VB.Menu mnuColorsGreen 
         Caption         =   "&Green"
      End
      Begin VB.Menu mnuColorsBlue 
         Caption         =   "&Blue"
      End
      Begin VB.Menu mnuColorsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorsBlack 
         Caption         =   "B&lack"
      End
      Begin VB.Menu mnuColorsWhite 
         Caption         =   "&White"
      End
      Begin VB.Menu mnuColorsSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorsQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmQBColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmQBColors.frm
'
' Project: Colors
'
' Purpose: Demostrate the use of MDI Form, Menus as well as Color functions
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuColorsBlack_Click()
    BackColor = QBColor(0) ' Black
End Sub

Private Sub mnuColorsBlue_Click()
    BackColor = QBColor(1) ' Blue
End Sub

Private Sub mnuColorsGreen_Click()
    BackColor = QBColor(2) ' Green
End Sub

Private Sub mnuColorsQuit_Click()
    Unload Me
End Sub

Private Sub mnuColorsRed_Click()
    BackColor = QBColor(4) 'Red
End Sub

Private Sub mnuColorsWhite_Click()
   BackColor = QBColor(15) 'white
End Sub
