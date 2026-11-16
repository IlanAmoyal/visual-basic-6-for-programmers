VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4935
   ClientLeft      =   8895
   ClientTop       =   6180
   ClientWidth     =   4980
   LinkTopic       =   "Form3"
   ScaleHeight     =   4935
   ScaleWidth      =   4980
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: Form3.frm
'
' Project: TraceEvents
'
' Purpose: Demonstrate the squence of the events - Initialize, Load,
'          Unload , QueryUnload, GotFocus, LostFocus, Activate, Deactivate
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Private Sub Form_Activate()
    Form1.Print "Form3: Activate"
End Sub

Private Sub Form_Deactivate()
    Form1.Print "Form3: Deactivate"
End Sub

Private Sub Form_GotFocus()
    Form1.Print "Form3: GotFocus"
End Sub

Private Sub Form_Initialize()
    Form1.Print "Form3: Initialize"
End Sub

Private Sub Form_Load()
    Form1.Print "Form3: Load"
End Sub

Private Sub Form_LostFocus()
    Form1.Print "Form3: LostFocus"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form1.Print "Form3: QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Print "Form3: Unload"
End Sub

