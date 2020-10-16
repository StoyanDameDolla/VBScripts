Option Explicit 

Dim password 

Do
	password = InputBox("Please, type in your password in here...", "Verification stage")
	If IsEmpty(password) Then
		WScript.Quit
	ElseIf password = "" Then
		MsgBox "Please, do not leave a blank field..."
	ElseIf password = "basketball" Then
		MsgBox "Password is correct!"
		Exit Do
	Else
		MsgBox "Re-enter the correct password!"
	End If
Loop