Option Explicit

Dim firstName

firstName = InputBox("What is your name?", "Information:")

If firstName = "Stoyan" Then
	MsgBox "Hello, " + firstName
ElseIf firstName <> "Stoyan" Then
	MsgBox "This is not the right name!"
End If