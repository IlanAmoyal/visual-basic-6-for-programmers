VERSION 5.00
Begin VB.Form frmProcessInputs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Inputs"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmProcessInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmProcessInputs
'
' Purpose: Demostrate the use of Do...Loop
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim nCurrVal As Long, nMul As Long, nSum As Long, nCounter As Long
    nCounter = 0
    nSum = 0
    nMul = 1
    nCurrVal = Val(InputBox("Enter an integer:"))
    Do Until nCurrVal = 0
        nCounter = nCounter + 1
        nSum = nSum + nCurrVal
        nMul = nMul * nCurrVal
        nCurrVal = Val(InputBox("Enter an integer:"))
    Loop
    If nCounter > 0 Then
        Cls
        Print Str(nCounter); " values were enterd"
        Print "Statistics:"
        Print vbTab & "Sum: " & Str(nSum)
        Print vbTab & "Mul: " & Str(nMul)
        Print vbTab & "Avg: " & Str(nSum / nCounter)
    Else
        Print "No values were enterd"
    End If
End Sub

