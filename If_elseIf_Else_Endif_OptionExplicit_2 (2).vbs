Option Explicit

Dim a 

a = MsgBox("Guess a button", vbAbortRetryIgnore)

If a = vbRetry Then
	MsgBox "Correct", vbInformation
ElseIf a = vbIgnore Then
	MsgBox "Wrong", vbExclamation
Else
	MsgBox "Command retried", vbExclamation
End If