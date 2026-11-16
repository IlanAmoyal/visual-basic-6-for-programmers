VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workers"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "cmdPrint"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtReport 
      Height          =   4575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmMain.frx":0000
      Top             =   3000
      Width           =   6615
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Text            =   "txtPhone"
      Top             =   90
      Width           =   2655
   End
   Begin VB.TextBox txtAddress 
      Height          =   765
      Left            =   840
      TabIndex        =   7
      Text            =   "txtAddress"
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "txtName"
      Top             =   480
      Width           =   5895
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "txtID"
      Top             =   90
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "cmdAdd"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "A&ddress:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "P&hone"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   180
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&ID:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   210
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   120
      X2              =   6840
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const SIZE_WORKERS As Integer = 20
Private g_nWorkers As Integer
Dim Workers(0 To SIZE_WORKERS - 1) As Worker

Private Sub cmdAdd_Click()
    If g_nWorkers >= SIZE_WORKERS Then
        MsgBox "The array is full."
        Exit Sub
    End If
    SetWorkerData Workers(g_nWorkers), _
                  Val(txtId.Text), _
                  txtName.Text, _
                  txtAddress.Text, _
                  txtPhone.Text
    g_nWorkers = g_nWorkers + 1
    ClearAllText
    txtId.SetFocus
End Sub

Private Sub cmdPrint_Click()
    txtReport.Text = ""
    PrintWorkerData txtReport, Workers, g_nWorkers
End Sub

Private Sub Form_Load()
    ClearAllText
    txtReport.Text = ""
    cmdAdd.Caption = "&Add"
    cmdPrint.Caption = "&Print"
    g_nWorkers = 0
End Sub

Private Sub txtAddress_GotFocus()
    SelAllText txtAddress
End Sub

Private Sub txtId_GotFocus()
    SelAllText txtId
End Sub

Private Sub SelAllText(txt As TextBox)
    txt.SelStart = 1
    txt.SelLength = Len(txt.Text)
End Sub

Private Sub txtName_GotFocus()
    SelAllText txtName
End Sub

Private Sub txtPhone_GotFocus()
    SelAllText txtPhone
End Sub

Private Sub ClearAllText()
    txtAddress.Text = ""
    txtId.Text = ""
    txtName.Text = ""
    txtPhone.Text = ""
End Sub
