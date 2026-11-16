VERSION 5.00
Begin VB.Form frmRGB 
   Caption         =   "RGB"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
      Begin VB.Menu mnuColorsCyan 
         Caption         =   "&Cyan"
      End
      Begin VB.Menu mnuColorMagenta 
         Caption         =   "&Magenta"
      End
      Begin VB.Menu mnuColorsYellow 
         Caption         =   "&Yellow"
      End
      Begin VB.Menu mnuColorsSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorsQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmRGB.frm
'
' Project: Colors
'
' Purpose: Demostrate the use of MDI Form, Menus as well as Color functions
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuColorMagenta_Click()
    BackColor = RGB(255, 0, 255) 'Magenta
End Sub

Private Sub mnuColorsBlue_Click()
    BackColor = RGB(0, 0, 255) ' Blue
End Sub

Private Sub mnuColorsCyan_Click()
    BackColor = RGB(0, 255, 255) ' Cyan
End Sub

Private Sub mnuColorsGreen_Click()
    BackColor = RGB(0, 255, 0) ' Green
End Sub

Private Sub mnuColorsQuit_Click()
    Unload Me
End Sub

Private Sub mnuColorsRed_Click()
    BackColor = RGB(255, 0, 0) ' Red
End Sub

Private Sub mnuColorsYellow_Click()
    BackColor = RGB(255, 255, 0) ' Yellow
End Sub
