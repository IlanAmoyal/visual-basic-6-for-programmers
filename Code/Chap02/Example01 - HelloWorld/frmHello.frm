VERSION 5.00
Begin VB.Form frmHello 
   AutoRedraw      =   -1  'True
   Caption         =   "Hello World Program"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name: frmHello.frm
'
' Project: HelloWorld
'
' Purpose: Demonstrate the use of Print, MsgBox and autoRedraw
'
' Copyright (C) 2001 Ilan Amoyal, All rights reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Click()
    Print "Hello, world!" 'Print the message on the form
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    MsgBox "Hello, world!" 'Display a message box to the user
End Sub
