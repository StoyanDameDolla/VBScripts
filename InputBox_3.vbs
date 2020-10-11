Option Explicit

Dim firstName

firstName = InputBox("What is your age?", "Information:")

If isNumeric(firstName) And firstName = "15" Then
MsgBox "Correct"
ElseIf firstName <> "15" Then
MsgBox "Wrong age input!"
ElseIf Not isNumeric(firstName) Then
MsgBox "Please, type in a numeric value!"
End If