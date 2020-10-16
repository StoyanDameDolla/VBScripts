Option Explicit

Dim first_name, last_name, full_name

first_name = InputBox("Firstname: ")
last_name = InputBox("Lastname: ")

Call AddFunction(first_name, last_name)

Sub AddFunction(value1, value2)
	Dim full_name
	full_name = value1 & " " & value2
	MsgBox "Your full name is: " & full_name     'Sub cannot be set to equal the end variable value.
End Sub