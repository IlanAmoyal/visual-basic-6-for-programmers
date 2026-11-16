VERSION 5.00
Begin VB.Form frmStudentCourses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Courses"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUIUpdate 
      Interval        =   20
      Left            =   6960
      Top             =   2040
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "cmdQuit"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CheckBox chkDeleteAll 
      Caption         =   "chkDeleteAll"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "cmdDelete"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   4095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "cmdAdd"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.ComboBox cboCourses 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox cboStudents 
      Height          =   315
      ItemData        =   "frmStudentCourses.frx":0000
      Left            =   240
      List            =   "frmStudentCourses.frx":0002
      TabIndex        =   1
      Text            =   "cboStudents"
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox lstStudentToCourse 
      Height          =   1410
      ItemData        =   "frmStudentCourses.frx":0004
      Left            =   3360
      List            =   "frmStudentCourses.frx":000B
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   480
      Width           =   4095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   3000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   3000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   3120
      X2              =   3120
      Y1              =   360
      Y2              =   3120
   End
   Begin VB.Label lblCourses 
      AutoSize        =   -1  'True
      Caption         =   "lblCourses"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblStudentToCourse 
      AutoSize        =   -1  'True
      Caption         =   "lblStudentToCourse"
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblstudents 
      AutoSize        =   -1  'True
      Caption         =   "lblStudents"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   3120
      X2              =   3120
      Y1              =   360
      Y2              =   3120
   End
End
Attribute VB_Name = "frmStudentCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: frmStudentCourses
'
' Purpose: Demonstrate the use of CheckBox, ListBox and ComboBox
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Dim m_nCheckCount As Integer ' Holds the number of checked items in
                             ' lstStudentToCourse ListBox
Dim m_nCheckedSource As Integer ' Hold a value that determines whether
                                ' a ItemCheck event (of lstStudentToCourse)
                                ' was invoked because of the code placed
                                ' in the Click Event of chkDeleteAll
                                ' or because of the user activity

Const CHECK_SOURCE_USER As Integer = 0 ' Means the code Invoked the event
Const CHECK_SOURCE_CODE As Integer = 1 ' Means the user Invoked the event

Private Sub chkDeleteAll_Click()
    Dim i As Integer
    Dim bChecked As Boolean
    
    If m_nCheckedSource = CHECK_SOURCE_CODE Then Exit Sub
    
    bChecked = (chkDeleteAll.Value = vbChecked)
    m_nCheckedSource = CHECK_SOURCE_CODE
    
    For i = 0 To lstStudentToCourse.ListCount - 1
        lstStudentToCourse.Selected(i) = bChecked
    Next i
    
    If bChecked Then
        m_nCheckCount = lstStudentToCourse.ListCount
    Else
        m_nCheckCount = 0
    End If
    
    m_nCheckedSource = CHECK_SOURCE_USER
End Sub

Private Sub cmdAdd_Click()
    Dim strRow As String
    strRow = cboStudents.Text & ": " & cboCourses.Text
    lstStudentToCourse.AddItem strRow
    cboStudents.SelStart = 0
    cboStudents.SelLength = Len(cboStudents.Text)
    cboStudents.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    ' Check whether to remove all items
    If chkDeleteAll.Value = vbChecked Then
        lstStudentToCourse.Clear ' Use Clear to remove all items
    Else
        ' Remove only the checked items
        i = 0
        Do While i <= (lstStudentToCourse.ListCount - 1)
            If lstStudentToCourse.Selected(i) = True Then
                lstStudentToCourse.RemoveItem (i)
            Else
                i = i + 1
            End If
        Loop
    End If
    
    ' Reset the "Delete all" - CheckBox and the
    ' Checked-Items counter (m_nCheckCount)
    chkDeleteAll.Value = vbUnchecked
    m_nCheckCount = 0
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Initializing lblStudets
    lblstudents.Caption = "&Students"
    
    ' Initializing cboStudents
    cboStudents.Clear
    cboStudents.AddItem "John Degweed"
    cboStudents.AddItem "Erick E"
    cboStudents.AddItem "Michel de Hey"
    cboStudents.AddItem "Paul Okenfold"
    cboStudents.AddItem "Dimmitri"
    cboStudents.AddItem "Tiesto"
    cboStudents.AddItem "Josh Wink"
    cboStudents.AddItem "Terry Francis"
    cboStudents.AddItem "Ferry Corsten"
    
    ' Initializing lblCourses
    lblCourses.Caption = "&Courses"
    
    ' Initializing cboCourses
    cboCourses.Clear
    cboCourses.AddItem "Advanced DJ Techniqs"
    cboCourses.AddItem "Using guitar loops in the Mix"
    cboCourses.AddItem "Synthisizer - Advance Topics"
    cboCourses.AddItem "Synthisizer - Fundamentals"
    cboCourses.AddItem "Extarcting Loops"
    
    ' Initializing cmdAdd
    cmdAdd.Caption = "&Add to list"
    
    ' Initializing cmdQuit
    cmdQuit.Caption = "&Quit"
    
    ' Initializing lblStudentsToCourses
    lblStudentToCourse.Caption = "Student to course"
    
    ' Initializing lstStudentToCourse
    lstStudentToCourse.Clear
    
    ' Initializing chkDeleteAll
    chkDeleteAll.Caption = "Delete &all rows"
    
    ' Initializing cmdDelete
    cmdDelete.Caption = "&Delete"
    
    ' Other initializations
    m_nCheckCount = 0
    m_nCheckedSource = CHECK_SOURCE_USER
End Sub

Private Sub lstStudentToCourse_ItemCheck(Item As Integer)
    Dim nCheckedCount As Integer, i As Integer
    
    ' Process this code only if the Checked or Unchecked operation
    ' was invoked by the user and not the code inside chkDeleteAll
    If m_nCheckedSource = CHECK_SOURCE_USER Then
        m_nCheckCount = m_nCheckCount + _
                        IIf(lstStudentToCourse.Selected(Item), 1, -1)
        m_nCheckedSource = CHECK_SOURCE_CODE
        If m_nCheckCount = 0 Then
            chkDeleteAll.Value = vbUnchecked
        ElseIf m_nCheckCount = lstStudentToCourse.ListCount Then
            chkDeleteAll.Value = vbChecked
        Else
            chkDeleteAll.Value = vbGrayed
        End If
        m_nCheckedSource = CHECK_SOURCE_USER
    End If
End Sub

Private Sub tmrUIUpdate_Timer()
    ' Enable the "Add to list" command only if there is text in
    ' cboStudents as well as in cboCourses
    cmdAdd.Enabled = cboCourses.Text <> "" And cboStudents.Text <> ""
    ' Enable the "Delete All" - CheckBox  only if there are
    ' items in the list
    chkDeleteAll.Enabled = lstStudentToCourse.ListCount > 0
    ' Enable cmdDelete only if the number of the checked items
    ' is greater than 0
    cmdDelete.Enabled = m_nCheckCount > 0
End Sub
