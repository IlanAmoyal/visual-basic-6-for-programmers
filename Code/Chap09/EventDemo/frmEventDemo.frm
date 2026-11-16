VERSION 5.00
Begin VB.Form frmEventDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Event Demo"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetAge 
      Caption         =   "Set Age"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSetName 
      Caption         =   "Set Name"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEventDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents prs As CPerson
Attribute prs.VB_VarHelpID = -1

Private Sub cmdReset_Click()
    prs.Reset
End Sub

Private Sub cmdSetAge_Click()
    prs.Age = Val(InputBox("Enter the age"))
End Sub

Private Sub cmdSetName_Click()
    prs.Name = InputBox("Enter the name")
End Sub

Private Sub Form_Load()
    Set prs = New CPerson
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set prs = Nothing
End Sub

Private Sub prs_Changed()
    MsgBox "prs has changed. The new values are:" & vbCrLf & _
           "Name = " & prs.Name & vbCrLf & _
           "Age = " & Trim(Str(prs.Age))
End Sub

Private Sub prs_WillReset(bCancel As Boolean)
    If MsgBox("Do you want to cancel the reset?", vbYesNo) = vbYes Then
        bCancel = True
    End If
End Sub
