VERSION 5.00
Begin VB.Form frmCOMPlusClient
   Caption         =   "COM+ Transaction Demo"
   ClientHeight    =   1515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFail
      Caption         =   "Simulate failure"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdRun
      Caption         =   "Run transaction"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmCOMPlusClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRun_Click()
	Dim obj As Object

	On Error GoTo Failed

	Set obj = CreateObject("MTSSampleDLL.MTSObj")
	Call obj.Transfer(chkFail.Value = vbChecked)

	MsgBox "Transaction completed"
	Exit Sub

Failed:
	MsgBox Err.Description
End Sub
