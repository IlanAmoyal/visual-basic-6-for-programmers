VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   8865
   ClientTop       =   960
   ClientWidth     =   5010
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   5010
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: Form2.frm
'
' Project: TraceEvents
'
' Purpose: Demonstrate the squence of the events - Initialize, Load,
'          Unload , QueryUnload, GotFocus, LostFocus, Activate, Deactivate
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Private Sub Form_Activate()
    Form1.Print "Form2: Activate"
End Sub

Private Sub Form_Deactivate()
    Form1.Print "Form2: Deactivate"
End Sub

Private Sub Form_GotFocus()
    Form1.Print "Form2: GotFocus"
End Sub

Private Sub Form_Initialize()
    Form1.Print "Form2: Initialize"
End Sub

Private Sub Form_Load()
    Form1.Print "Form2: Load"
End Sub

Private Sub Form_LostFocus()
    Form1.Print "Form2: LostFocus"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form1.Print "Form2: QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Print "Form2: Unload"
End Sub
