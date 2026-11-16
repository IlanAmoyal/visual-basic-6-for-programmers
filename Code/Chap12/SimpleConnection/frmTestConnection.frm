VERSION 5.00
Begin VB.Form frmTestConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Connection"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmTestConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As ADODB.Connection

Private Sub Form_Load()
    Set conn = New ADODB.Connection
End Sub

Private Sub cmdConnect_Click()
    conn.ConnectionString = "DSN=NwindDSN;User ID=sa"
    conn.Open
    If conn.State = adStateOpen Then
        MsgBox "Connection successful"
        cmdConnect.Enabled = False
        cmdDisconnect.Enabled = True
    Else
        
        MsgBox "Connection Failed"
    End If
End Sub

Private Sub cmdDisconnect_Click()
    conn.Close
    If conn.State = adStateClosed Then
        MsgBox "Disconnected"
        cmdConnect.Enabled = True
        cmdDisconnect.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If conn.State = adStateOpen Then
        conn.Close
    End If
    Set conn = Nothing
End Sub
