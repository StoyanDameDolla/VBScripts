Option Explicit

MsgBox("Fullname is: " & AddFunction(InputBox("Firstname: "), InputBox("Lastname: ")))

MsgBox("Hello, " & AddFunction(InputBox("Firstname: "), InputBox("Lastname: ")))

'Two iterations will be made - first MsgBox with First and last name + Fullname prompt, second MsgBox with Hello, first + last name.
Function AddFunction(first_name, last_name)
	Dim full_name
	full_name = first_name & " " & last_name
	AddFunction = full_name
End Function