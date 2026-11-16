VERSION 5.00
Begin VB.Form frmSystemColors 
   Caption         =   "System Colors"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Menu mnuColors 
      Caption         =   "&Colors"
      Begin VB.Menu mnuColorsButtonFace 
         Caption         =   "&Button Face"
      End
      Begin VB.Menu mnuColorsToolTip 
         Caption         =   "&Tool Tip"
      End
      Begin VB.Menu mnuColorsDesktop 
         Caption         =   "&Desktop"
      End
      Begin VB.Menu mnuColorsSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorsQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmSystemColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmSystemColors.frm
'
' Project: Colors
'
' Purpose: Demostrate the use of MDI Form, Menus as well as Color functions
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuColorsButtonFace_Click()
    BackColor = SystemColorConstants.vbButtonFace
End Sub

Private Sub mnuColorsDesktop_Click()
    BackColor = SystemColorConstants.vbDesktop
End Sub

Private Sub mnuColorsQuit_Click()
    Unload Me
End Sub

Private Sub mnuColorsToolTip_Click()
    BackColor = SystemColorConstants.vbInfoBackground ' ToolTip background
End Sub
