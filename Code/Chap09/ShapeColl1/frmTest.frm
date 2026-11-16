VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shape Collection Demo 1"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdListItemsMethod2 
      Caption         =   "List items - Method 2"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtColl 
      BackColor       =   &H8000000F&
      Height          =   3735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton cmdListItemsMethod1 
      Caption         =   "List items - Method 1"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddCircle 
      Caption         =   "Add Circle"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddRect 
      Caption         =   "Add Rectangle"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public shpColl As Collection

Private Sub Form_Load()
    Set shpColl = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set shpColl = Nothing
End Sub

Private Sub cmdAddCircle_Click()
    Dim cr As CCircle
    Set cr = New CCircle
    cr.Radius = Val(InputBox("Enter the radius for the circle:"))
    shpColl.Add cr
    Set cr = Nothing
End Sub

Private Sub cmdAddRect_Click()
    Dim rc As CRect
    Set rc = New CRect
    rc.Length = Val(InputBox("Enter the length of the rectangle:"))
    rc.Width = Val(InputBox("Enter the width of the rectangle:"))
    shpColl.Add rc
    Set rc = Nothing
End Sub

Private Sub cmdListItemsMethod1_Click()
    Dim obj As Object
    txtColl.Text = ""
    On Error Resume Next
    For Each obj In shpColl
        txtColl.Text = txtColl.Text & TypeName(obj) & vbCrLf
        txtColl.Text = txtColl.Text & String(8, " ") & _
                       "Length = " & obj.Length & vbCrLf
        txtColl.Text = txtColl.Text & String(8, " ") & _
                       "Width = " & obj.Width & vbCrLf
        txtColl.Text = txtColl.Text & String(8, " ") & _
                       "Radius = " & obj.Radius & vbCrLf
        txtColl.Text = txtColl.Text & String(8, " ") & _
                       "Surface = " & obj.Surface & vbCrLf
        txtColl.Text = txtColl.Text & String(8, " ") & _
                       "Perimeter = " & obj.Perimeter & vbCrLf
    Next
End Sub

Private Sub cmdListItemsMethod2_Click()
    Dim obj As Object
    Dim rc As CRect
    Dim cr As CCircle
    txtColl.Text = ""
    On Error Resume Next
    For Each obj In shpColl
        If TypeOf obj Is CRect Then
            Set rc = obj
            txtColl.Text = txtColl.Text & "CRect" & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Length = " & rc.Length & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Width = " & rc.Width & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Surface = " & rc.Surface & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Perimeter = " & rc.Perimeter & vbCrLf
        Else
            Set cr = obj
            txtColl.Text = txtColl.Text & "CCircle" & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Radius = " & cr.Radius & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Surface = " & rc.Surface & vbCrLf
            txtColl.Text = txtColl.Text & String(8, " ") & _
                           "Perimeter = " & rc.Perimeter & vbCrLf
        End If
    Next
End Sub
