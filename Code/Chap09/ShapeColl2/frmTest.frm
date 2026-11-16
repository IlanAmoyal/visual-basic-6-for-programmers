VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shapes Collection Demo 2"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddRect 
      Caption         =   "Add Rectangle"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddCircle 
      Caption         =   "Add Circle"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdListItems 
      Caption         =   "List items"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtColl 
      BackColor       =   &H8000000F&
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public shpColl As Shapes

Private Sub Form_Load()
    Set shpColl = New Shapes
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

Private Sub cmdListItems_Click()
    Dim shp As IShape
    txtColl.Text = ""
    For Each shp In shpColl
        txtColl.Text = txtColl.Text & shp.FormatDetails
    Next
End Sub
