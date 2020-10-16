Option Explicit

Dim first_name, last_name

first_name = InputBox("Firstname: ")
last_name = InputBox("Lastname: ")

MsgBox "Fullname is: " & AddFunction(first_name, last_name)

Function AddFunction(value1, value2)
	Dim full_name
	full_name = value1 & " " & value2
	AddFunction = full_name
End Function