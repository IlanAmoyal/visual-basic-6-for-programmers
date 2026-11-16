VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   3480
   ClientTop       =   1215
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   5115
   Begin VB.Menu mnuForm 
      Caption         =   "&Form"
      Begin VB.Menu mnuForm2 
         Caption         =   "2"
      End
      Begin VB.Menu mnuForm3 
         Caption         =   "3"
      End
      Begin VB.Menu mnuFormSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuFormExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: Form1.frm
'
' Project: TraceEvents
'
' Purpose: Demonstrate the squence of the events - Initialize, Load,
'          Unload , QueryUnload, GotFocus, LostFocus, Activate, Deactivate
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuForm2_Click()
    ' Display Form2
    Load Form2
    Form2.Show
End Sub

Private Sub mnuForm3_Click()
    ' Display Form3
    Load Form3
    Form3.Show
End Sub

Private Sub mnuFormClear_Click()
    ' Clear the content of the form
    Cls
End Sub

Private Sub mnuFormExit_Click()
    ' Close the program
    Unload Form3
    Unload Form2
    Unload Me
End Sub
