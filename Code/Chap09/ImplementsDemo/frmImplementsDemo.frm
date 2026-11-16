VERSION 5.00
Begin VB.Form frmImplementsDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Implements Demo"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetShapeIsRect 
      Caption         =   "Set shape = r"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdSetShapeIsCircle 
      Caption         =   "Set shape = c"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtInfo 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   4815
   End
   Begin VB.CommandButton cmdDisplayInfo 
      Caption         =   "Display Info..."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton cmdSetRectangleValues 
      Caption         =   "Set Rectangle Values..."
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdSetCircle 
      Caption         =   "Set Circle Radius..."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmImplementsDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim c As New CCircle
Dim r As New CRect
Dim shape As IShape
Implements Hyperlink

Private Sub cmdSetCircle_Click()
    c.m_dRadius = Val(InputBox("Enter the radius of the circle:"))
End Sub

Private Sub cmdSetRectangleValues_Click()
    r.m_dLength = Val(InputBox("Enter the length of the recangle:"))
    r.m_dWidth = Val(InputBox("Enter the width of the rectangle:"))
End Sub

Private Sub cmdSetShapeIsCircle_Click()
    Set shape = c
End Sub

Private Sub cmdSetShapeIsRect_Click()
    Set shape = r
End Sub

Private Sub cmdDisplayInfo_Click()
    If shape Is Nothing Then Exit Sub
    txtInfo.Text = TypeName(shape) & "..." & vbCrLf
    txtInfo.Text = txtInfo.Text & String(8, " ") & _
                   "Perimeter = " & Str(shape.Perimeter()) & vbCrLf
    txtInfo.Text = txtInfo.Text & String(8, " ") & _
                   "Surface = " & Str(shape.Surface()) & vbCrLf

End Sub
