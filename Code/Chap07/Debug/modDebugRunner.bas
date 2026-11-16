Attribute VB_Name = "modDebugRunner"
Option Explicit

Public Sub Main()
    Dim inputValue As Double

    inputValue = Val(InputBox("Enter an integer from 0 to 255:", "Debug example"))

    If inputValue < 0 Or inputValue > 255 Or inputValue <> Fix(inputValue) Then
        MsgBox "Enter a whole number from 0 to 255.", vbExclamation
        Exit Sub
    End If

    MsgBox "SumUpTo(" & CStr(CInt(inputValue)) & ") = " & _
           CStr(SumUpTo(CInt(inputValue)))
End Sub
