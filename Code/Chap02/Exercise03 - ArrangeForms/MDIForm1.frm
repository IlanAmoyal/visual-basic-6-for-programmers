VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8085
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11310
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuShow 
      Caption         =   "&Show"
      Begin VB.Menu mnuShowForm1 
         Caption         =   "Form&1"
      End
      Begin VB.Menu mnuShowForm2 
         Caption         =   "Form&2"
      End
      Begin VB.Menu mnuShowForm3 
         Caption         =   "Form&3"
      End
      Begin VB.Menu mnuShowSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindowTileVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: MDIForm1.frm
'
' Project: ArrangeForms
'
' Purpose: Demostrate the use of MDI Form, Menus as well as Color functions
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuShowExit_Click()
    Unload Me
End Sub

Private Sub mnuShowForm1_Click()
    Load Form1
    Form1.Show
End Sub

Private Sub mnuShowForm2_Click()
    Load Form2
    Form2.Show
End Sub

Private Sub mnuShowForm3_Click()
    Load Form3
    Form3.Show
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
