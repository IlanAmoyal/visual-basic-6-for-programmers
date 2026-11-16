VERSION 5.00
Begin VB.Form frmTestProcs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmTestProcs"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   1815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAvg 
      Caption         =   "cmdAvg"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSumCalls 
      Caption         =   "cmdSumCalls"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGiveMe5 
      Caption         =   "cmdGiveMe5"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindBigger 
      Caption         =   "cmdFindBigger"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "cmdFactorial"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "cmdSwap"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdDispSum 
      Caption         =   "cmddispSum"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdHello 
      Caption         =   "cmdHello"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmTestProcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmTestProcs
'
' Purpose: This form tests the procedures written in modProcs
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Private Sub cmdAvg_Click()
    ' Test the avarage function procedure
    MsgBox "The average of 9, 10, 0, 6 and 4 is: " & vbCrLf & _
           Trim(Str(Avg(9, 10, 0, 6, 4)))
End Sub

Private Sub cmdDispSum_Click()
    ' Collect values from the user
    Dim a As Integer, b As Integer
    a = Val(InputBox("enter a value for a:"))
    b = Val(InputBox("Enter a value for b:"))
    ' Call DispSum
    DispSum a, b
End Sub

Private Sub cmdFactorial_Click()
    Dim i As Integer
    ' Ask for input from the user
    i = Val(InputBox("enter an integer (not too greater!):"))
    ' Call Fuctorial and displays the result
    MsgBox "The factorial of " & Trim(Str(i)) & " is:" & vbCrLf & _
           Str(Factorial(i))
End Sub

Private Sub cmdFindBigger_Click()
    ' Ask for values from the user
    Dim arr(0 To 4) As Integer
    Dim i As Integer, nBigger As Integer
    Dim strMsg As String
    ' Ask the user for inputs
    strMsg = ""
    For i = 0 To 4
        arr(i) = Val(InputBox("Enter a value (" & Trim(Str(5 - i)) & " elements rmain:)"))
        strMsg = strMsg & "arr(" & Trim(Str(i)) & ") = " & Trim(Str(arr(i))) & vbCrLf
    Next i
    
    ' Call FindBigger
    nBigger = FindBigger(arr)
    
    ' Display the result
    MsgBox "The original array is: " & vbCrLf & _
           strMsg & _
           "The bigger value is: " & vbCrLf & _
           "arr(" & Trim(Str(nBigger)) & ") = " & Trim(Str(arr(nBigger)))
End Sub

Private Sub cmdGiveMe5_Click()
    ' Call GiveMe5 and displays the result
    MsgBox "The value returned by GiveMe5 is: " & Trim(Str(GiveMe5()))
End Sub

Private Sub cmdHello_Click()
    ' Call Hello sub procedure
    Call Hello
End Sub

Private Sub cmdSumCalls_Click()
    ' Call SumCalls
    MsgBox "The function SumCalls was called " & Trim(Str(SumCalls())) & " times"
End Sub

Private Sub cmdSwap_Click()
    ' Collect values from the user
    Dim a As Variant, b As Variant
    a = InputBox("Enter a string for var1:")
    b = InputBox("Enter a string for var2:")
    ' Display the values before the swap
    MsgBox "Before swaping: " & vbCrLf & _
           "a = " & a & vbCrLf & _
           "b = " & b
    ' Call swap
    Swap a, b
    'Display the values after the swaping
    MsgBox "After swaping: " & vbCrLf & _
           "a = " & a & vbCrLf & _
           "b = " & b
End Sub

Private Sub Form_Load()
    ' Initializations
    Me.Caption = "Test"
    cmdHello.Caption = "Hello"
    cmdDispSum.Caption = "DispSum"
    cmdSwap.Caption = "Swap"
    cmdGiveMe5.Caption = "GiveMe5"
    cmdFactorial.Caption = "Factorial"
    cmdFindBigger.Caption = "FindBigger"
    cmdSumCalls.Caption = "SumCalls"
    cmdAvg.Caption = "Avg"
End Sub

