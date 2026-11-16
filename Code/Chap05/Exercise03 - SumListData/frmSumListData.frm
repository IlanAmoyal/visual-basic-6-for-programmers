VERSION 5.00
Begin VB.Form frmSumListData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmSumListData"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "cmdQuit"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSum 
      Caption         =   "cmdSum"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "cmdInsert"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ListBox lstItems 
      Height          =   2010
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtAdditionalData 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Text            =   "txtAdditionalData"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "txtName"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblAdditionalData 
      AutoSize        =   -1  'True
      Caption         =   "lblAdditionalData"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "lblName"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "frmSumListData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmSumListData
'
' Purpose: Demonstrate the use of ListBox
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub cmdInsert_Click()
    Dim lItemData As Long
    If txtName.Text <> "" Then
        lstItems.AddItem txtName.Text
        lItemData = IIf(txtAdditionalData.Text = "", 1&, Val(txtAdditionalData.Text))
        lstItems.ItemData(lstItems.NewIndex) = lItemData
        txtName.SetFocus
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSum_Click()
    Dim lSum As Long, nIndex As Integer
    lSum = 0
    For nIndex = 0 To lstItems.ListCount - 1
        lSum = lSum + lstItems.ItemData(nIndex)
    Next
    MsgBox "The total sum is: " & Str(lSum)
End Sub

Private Sub Form_Load()
    lblAdditionalData.Caption = "Additional Data"
    lblName.Caption = "&Name"
    txtName.Text = ""
    txtAdditionalData.Text = ""
    lstItems.Clear
    cmdInsert.Caption = "&Insert"
    cmdQuit.Caption = "&Quit"
    cmdSum.Caption = "&Sum"
End Sub

Private Sub txtAdditionalData_GotFocus()
    txtAdditionalData.SelStart = 0
    txtAdditionalData.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub
