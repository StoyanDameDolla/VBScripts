Sub Greetings(username)
MsgBox "Hello, "& username & "!"
End sub

x = InputBox("Please, type in your name...")
Call Greetings(x)