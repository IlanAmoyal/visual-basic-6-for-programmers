VERSION 5.00
Begin VB.Form frmAnalyzer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Late Binding VS. Early Binding"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "I got it - STOP!"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtTimer 
      BackColor       =   &H8000000F&
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdUseObject 
      Caption         =   "Dim ... As Object"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSetNew 
      Caption         =   "Set ... = New"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSetNew_Click()
    Dim cls As CTest
    Dim lVal As Long, i As Integer
    Dim sTime As Single
    Set cls = New CTest
    sTime = Timer
    For i = 1 To 1000
        txtTimer.Text = "As CTest... " & vbCrLf & _
                        "i = " & Str(i) & vbCrLf & _
                        "Time elapsed:" & Str(Timer - sTime) & " sec"
        DoEvents
        For lVal = 1 To 10000
            cls.Val = lVal
        Next
    Next
End Sub

Private Sub cmdUseObject_Click()
    Dim cls As Object
    Dim lVal As Long, i As Integer
    Dim sTime As Single
    Set cls = New CTest
    sTime = Timer
    For i = 1 To 1000
        txtTimer.Text = "As Object..." & vbCrLf & _
                        "i = " & Str(i) & vbCrLf & _
                        "Time elapsed:" & Str(Timer - sTime) & " sec"
        DoEvents
        For lVal = 1 To 10000
            cls.Val = lVal
        Next
    Next
End Sub

Private Sub cmdStop_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
