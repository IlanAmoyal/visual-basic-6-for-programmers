Attribute VB_Name = "modProcs"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename: modProcs
'
' Purpose:  Demonstrate various ways to write sub procedures
'           and function procedures
'
' Copyright (C) 2001 by Ilan Amoyal, All rights Reserved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Example 01 - The Hello sub procedure
Public Sub Hello()
    ' Description:
    ' This sub procedure displays a message box
    ' contains the string "Hellom world!"
    MsgBox "Hello, world!"
End Sub

' Example 02 - The DispSum sub procedure
Public Sub DispSum(ByVal a As Integer, ByVal b As Integer)
    ' Description:
    ' This sub procedure displayes the sum of two numbers
    ' using message box
    MsgBox Str(a + b)
End Sub

' Example 03 - The Swap sub procedure
Public Sub Swap(ByRef var1 As Variant, ByRef var2 As Variant)
    ' Description:
    ' This sub procedure swaps between the content of two variables.
    ' Because the parameters were declared as Variant, any basic type
    ' of argument can be pass to this sub procedure
    Dim varTmp As Variant
    varTmp = var1
    var1 = var2
    var2 = varTmp
End Sub

' Example 04 - The GiveMe5 function procedure
Public Function GiveMe5() As Integer
    ' Description:
    ' This function procedure return 5 to the caller
    GiveMe5 = 5
End Function

' Example 05: The Factorial function procedure
Public Function Factorial(ByVal n As Byte) As Long
    ' Description:
    ' This function procedure get an integer and returns its factorial
    Dim lRes As Long
    lRes = 1
    Do While n > 0
        lRes = lRes * n
        n = n - 1
    Loop
    Factorial = lRes
End Function

' Example 06: The FindBigger function procedure
Public Function FindBigger(ByRef arr() As Integer) As Integer
    ' Description:
    ' This function search the given array to find the bigger element within it
    ' The return value is the index of the bigger element in the array
    Dim l As Integer, u As Integer
    Dim bigger As Integer
    
    l = LBound(arr)
    u = UBound(arr)
    
    bigger = l
    
    Do While l <= u
        If arr(l) > arr(bigger) Then bigger = l
        l = l + 1
    Loop
    FindBigger = bigger
End Function

' Example 07 - The SumCalls function procedure
Public Function SumCalls() As Integer
    ' Description:
    ' This function procedure returns the number of calls made
    ' to itself
    Static nCalls As Integer
    nCalls = nCalls + 1
    SumCalls = nCalls
End Function


'Example 08 - The Mult function procedure
Public Function Mult(ByVal Num1 As Double, _
                     ByVal Num2 As Double, _
                     ParamArray AdditionalNums()) As Double
    ' Description:
    ' This function multiply the arguments it got from the caller.
    ' The first and second parameters are not optional, but the number of the additional
    ' parameters can be vary
    Dim res As Double
    Dim l As Integer, u As Integer
    l = LBound(AdditionalNums)
    u = UBound(AdditionalNums)
    res = Num1 * Num2
    Do While l <= u
        res = res * AdditionalNums(l)
        l = l + 1
    Loop
    Mult = res
End Function


' Example 09 - The Div function procedure
Public Function Div(Optional ByVal Num1 As Double = 1#, _
                    Optional ByVal Num2 As Double = 1#) As Double
    ' Description:
    ' This function calculate the result of Num1 / Num2
    
    If Num2 = 0# Then Exit Function ' Num2 must be non-zero
    Div = Num1 / Num2
End Function





'Example 09 - The Avg function procedure
Public Function Avg(ParamArray numbers()) As Double
    ' Dsecription:
    ' This function calculates the average of the numbers it gets.
    Dim l As Integer, u As Integer, i As Integer, sum As Double
    
    l = LBound(numbers)
    u = UBound(numbers)
    If u < l Then
        ' If no numbers were passed to the function then exit function
        Avg = 0#
        Exit Function
    End If
    
    sum = 0
    For i = l To u
        sum = sum + Val(numbers(i))
    Next
    Avg = sum / (u - l + 1)
End Function
