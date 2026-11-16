Attribute VB_Name = "modDebug"
Option Explicit


Public Function SumUpTo(ByVal n As Integer) As Integer
	Dim res As Integer

	Debug.Print "The input for SumUpTo is : " & _
					Trim(Str(n)) ' Print to the immediate
	Debug.Assert n >= 0 ' Allow only non-negative arguments

	res = 0
	Do While n > 0
		res = res + n
		n = n - 1
	Loop
	SumUpTo = res
End Function
