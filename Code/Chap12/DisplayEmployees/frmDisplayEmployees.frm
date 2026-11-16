VERSION 5.00
Begin VB.Form frmDisplayEmployees 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Employees"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUseSQLStatement 
      Caption         =   "Use SQL Statement"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdUseTableName 
      Caption         =   "Use table name"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtDisplay 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   7215
   End
End
Attribute VB_Name = "frmDisplayEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As ADODB.Connection
Dim rcs As ADODB.Recordset

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.Open "Provider=SQLOLEDB.1;" & _
              "Persist Security Info=False;" & _
              "User ID=sa;" & _
              "Initial Catalog=Northwind;" & _
              "Data Source=WIN2K"
    If conn.State = adStateClosed Then
        MsgBox "Error: Fail to connect"
        Unload Me
    End If
    Set rcs = New ADODB.Recordset
    Set rcs.ActiveConnection = conn
End Sub

Private Sub cmdUseTableName_Click()
    If rcs.State = adStateOpen Then
        rcs.Close
    End If
    rcs.Source = "Employees"
    rcs.Open , , , , adCmdTable
    txtDisplay.Text = ""
    DisplayRecordset rcs
End Sub

Private Sub cmdUseSQLStatement_Click()
    If rcs.State = adStateOpen Then
        rcs.Close
    End If
    rcs.Open "SELECT * FROM Employees", , , , adCmdText
    txtDisplay.Text = ""
    DisplayRecordset rcs
End Sub

Private Sub DisplayRecordset(rs As ADODB.Recordset)
    rs.MoveFirst
    Do Until rs.EOF
        DisplaySingleRecord rs.Fields
        txtDisplay.Text = txtDisplay.Text & vbCrLf
        rs.MoveNext
    Loop
End Sub

Private Sub DisplaySingleRecord(flds As ADODB.Fields)
    Dim fld As ADODB.Field
    For Each fld In flds
        txtDisplay.Text = txtDisplay.Text & _
                          fld.Name & " = " & _
                          IIf(IsNull(fld.Value), "<Null>", fld.Value) & _
                          String(8, " ")
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rcs = Nothing
    Set conn = Nothing
End Sub

