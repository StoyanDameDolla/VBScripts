Option Explicit

Dim password, login

password = InputBox("Please, insert a password", "Password window")

Do until password = "basketball"
	Select case password
		Case "basketball"
			MsgBox "correct"
			login = 1
			Exit Do
		Case Else
			MsgBox "Please, provide a correct password!"
			password = InputBox("Please, insert a password", "Password window")
	End select
Loop

If login = 1 Then
WScript.Echo "Login successful!"
Else
WScript.Echo "Login unsuccessful!"
End if