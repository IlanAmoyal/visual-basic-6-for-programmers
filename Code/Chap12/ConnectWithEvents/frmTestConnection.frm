VERSION 5.00
Begin VB.Form frmTestConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Connection"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
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

Dim WithEvents conn As ADODB.Connection

Private Sub Form_Load()
    Set conn = New ADODB.Connection
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next
    conn.ConnectionString = "Provider=SQLOLEDB.1;" & _
                            "Persist Security Info=False;" & _
                            "User ID=sa;Initial Catalog=Northwind;" & _
                            "Data Source=WIN2K"
    conn.Open
    If Err.Number <> 0 Then
        MsgBox "Error No. " & Trim(Str(Err.Number)) & ":" & vbCrLf & _
               Err.Description
        Err.Clear
        conn.Errors.Clear
    End If
End Sub

Private Sub conn_WillConnect(ConnectionString As String, _
                             UserID As String, _
                             Password As String, _
                             Options As Long, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pConnection As ADODB.Connection)
    If adStatus = adStatusOK Then
        If MsgBox("Do you want to cancel the operation?", _
                  vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
            adStatus = adStatusCancel
        End If
    End If
End Sub

Private Sub conn_ConnectComplete(ByVal pError As ADODB.Error, _
                                 adStatus As ADODB.EventStatusEnum, _
                                 ByVal pConnection As ADODB.Connection)
    If conn.State = adStateOpen Then
        MsgBox "Connection successful"
        cmdConnect.Enabled = False
        cmdDisconnect.Enabled = True
    End If
End Sub

Private Sub cmdDisconnect_Click()
    conn.Close
End Sub

Private Sub conn_Disconnect(adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)
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
