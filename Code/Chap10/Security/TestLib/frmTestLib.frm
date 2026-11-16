VERSION 5.00
Begin VB.Form frmTestLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Login"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUsetGlobal 
      Caption         =   "Use Global"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdUseCreate 
      Caption         =   "Use ""CreateObject(...)"""
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdUseNew 
      Caption         =   "Use ""Set...New"""
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTestLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUseNew_Click()
    Dim sec As New SecurityLib.CLogin
    If sec.ShowLogin() Then
        MsgBox "You are succesfull logged in!"
    Else
        MsgBox "Login Failed!"
    End If
    Set sec = Nothing
End Sub

Private Sub cmdUseCreate_Click()
    Dim sec As Object
    Set sec = CreateObject("SecurityLib.CLogin")
    If sec.ShowLogin() Then
        MsgBox "You are succesfull logged in!"
    Else
        MsgBox "Login Failed!"
    End If
    Set sec = Nothing
End Sub

Private Sub cmdUsetGlobal_Click()
    If ShowLogin() Then
        MsgBox "You are succesfull logged in!"
    Else
        MsgBox "Login Failed!"
    End If
End Sub
