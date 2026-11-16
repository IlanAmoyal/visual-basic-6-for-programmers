VERSION 5.00
Begin VB.Form frmNavigateRecords 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Navigate Records"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "|< Move First"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "< Move Previous"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "Move Last >|"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "Move Next >"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
End
Attribute VB_Name = "frmNavigateRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As ADODB.Connection
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub cmdMoveFirst_Click()
    rs.MoveFirst
End Sub

Private Sub cmdMoveLast_Click()
    rs.MoveLast
End Sub

Private Sub cmdMoveNext_Click()
    rs.MoveNext
    
End Sub

Private Sub cmdMovePrevious_Click()
    rs.MovePrevious
    
End Sub

Private Sub Form_Load()
    Dim strSQLStatement As String
    
    ' Setting New Objects
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    ' Setting the connection
    conn.CommandTimeout = 30
    
    
    
'    conn.Provider = "SQLOLEDB.1"
'    conn.Properties("Persist Security Info") = False
'    conn.Properties("User ID") = "sa"
'    conn.Properties("Initial Catalog") = "Northwind"
    
    'conn.Properties("Data Source") = "WIN2K"
    
'    conn.Open "Provider=SQLOLEDB.1;" & _
'              "Persist Security Info=False;" & _
'              "User ID=sa;" & _
'              "Initial Catalog=Northwind;" & _
'              "Data Source=WIN2K"
    'conn.Mode = adModeRead
    conn.Open "File Name=" & App.Path & "\..\NWind.udl"
    
    
    If conn.State = adStateClosed Then
        MsgBox "Connection Failed!"
        Unload Me
    End If
    
    ' Preparing and executing SQL statement
    Set rs.ActiveConnection = conn
'    strSQLStatement = "SELECT emptbl.EmployeeID ID, " & _
'                      "emptbl.LastName [Last Name], " & _
'                      "emptbl.FirstName [First Name], " & _
'                      "mgrtbl.LastName  + ' ' + mgrtbl.FirstName Manager " & _
'                      "FROM Employees emptbl LEFT OUTER JOIN " & _
'                      "Employees mgrtbl ON emptbl.ReportsTo = mgrtbl.EmployeeID"
     strSQLStatement = "SELECT emptbl.EmployeeID As ID, " & _
                      "emptbl.LastName As [Last Name], " & _
                      "emptbl.FirstName As [First Name], " & _
                      "mgrtbl.LastName  + ' ' + mgrtbl.FirstName As Manager " & _
                      "FROM Employees emptbl LEFT OUTER JOIN " & _
                      "Employees mgrtbl ON emptbl.ReportsTo = mgrtbl.EmployeeID"
    rs.Open strSQLStatement
    rs.MoveFirst
End Sub

Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim fld As ADODB.Field
    Cls
    If rs.EOF Then
        rs.MoveLast
    ElseIf rs.BOF Then
        rs.MoveFirst
    Else
        For Each fld In rs.Fields
            Print fld.Name & ": " & IIf(IsNull(fld.Value), "<Null>", fld.Value)
        Next
    End If
End Sub


