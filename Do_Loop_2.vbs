Option Explicit

Dim pass

Do

pass = InputBox("Password")

If pass = "wired" then

exit do

ElseIf pass = "" then

MsgBox "Don't leave a blank field!"

ElseIf pass <> "wired" then

MsgBox "Incorrect input password!", vbCritical

End If

Loop

MsgBox "Correct!"
