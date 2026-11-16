VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 2  'RequiresTransaction
END
Attribute VB_Name = "MTSObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Transfer(ByVal Fail As Boolean)
	Dim cn As ADODB.Connection

	On Error GoTo Failed

	Set cn = New ADODB.Connection
	cn.Open "Provider=SQLOLEDB;Data Source=(local);" & _
		"Initial Catalog=COMPlusDemo;Integrated Security=SSPI;"


	cn.Execute "UPDATE Accounts SET Balance=Balance-10 WHERE Id=1"
	cn.Execute "UPDATE Accounts SET Balance=Balance+10 WHERE Id=2"

	If Fail Then Err.Raise vbObjectError + 1000, , "Test failure"

	cn.Close
	GetObjectContext.SetComplete
	Exit Sub

Failed:
	GetObjectContext.SetAbort
	Err.Raise vbObjectError + 1001, , "Transaction aborted"
End Sub
