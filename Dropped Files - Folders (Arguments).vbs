Set args = WScript.Arguments


If args.Count = 0 Then
MsgBox "Please, drop a file or folder."
Else
For each arg In args
	MsgBox arg
Next
End If
