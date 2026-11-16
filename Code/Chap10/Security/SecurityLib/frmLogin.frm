VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "&Username:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strUsername As String
Private m_strPassword As String

Private Sub cmdOK_Click()
    m_strUsername = txtUsername.Text
    m_strPassword = txtPassword.Text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Property Get Username() As String
    Username = m_strUsername
End Property

Public Property Get Password() As String
    Password = m_strPassword
End Property
