Option Explicit

MsgBox("Fullname is: " & AddFunction(InputBox("Firstname: "), InputBox("Lastname: ")))

MsgBox("The total value is " & AddFunction(InputBox("Fist value: "), InputBox("Second value: ")))

Function AddFunction(value1, value2)
	Dim end_value
	If IsNumeric(value1) And IsNumeric(value2) Then
		end_value = CInt(value1) + CInt(value2)		'Value are added.
		AddFunction = end_value
	Else
		end_value = value1 & " " & value2			'Values are concatenated.
		AddFunction = end_value
	End If
End Function