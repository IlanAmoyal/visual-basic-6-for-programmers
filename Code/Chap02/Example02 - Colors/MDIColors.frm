VERSION 5.00
Begin VB.MDIForm MDIColors 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11340
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuOpen 
      Caption         =   "&Open"
      Begin VB.Menu mnuOpenQBColorsForm 
         Caption         =   "&QBColors Form"
      End
      Begin VB.Menu mnuOpenRGBForm 
         Caption         =   "&RGB Form"
      End
      Begin VB.Menu mnuOpenSystemColorsForm 
         Caption         =   "&System Colors Form"
      End
      Begin VB.Menu mnuOpenSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MDIColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: mdiColors.frm
'
' Project: Colors
'
' Purpose: Demostrate the use of MDI Form, Menus as well as Color functions
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuOpenExit_Click()
    Unload Me
End Sub

Private Sub mnuOpenQBColorsForm_Click()
    ' Load and display frmQBColors
    Load frmQBColors
    frmQBColors.Show
End Sub

Private Sub mnuOpenRGBForm_Click()
    ' Load and display frmRGB
    Load frmRGB
    frmRGB.Show
End Sub

Private Sub mnuOpenSystemColorsForm_Click()
    ' Load and display frmSystemColors
    Load frmSystemColors
    frmSystemColors.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontally_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertically_Click()
    Arrange vbTileVertical
End Sub
