str = "Alex was born 04/12/1995. Julie was born 08/23/2000."

Set birth = New RegExp

birth.Pattern = "([A-Z][a-z]+)\swas\sborn\s((\d{2})/(\d{2})/(\d{4}))\b"
birth.Global = True

If birth.Test(str) Then
	For each item In birth.Execute(str)
		For Each submatch In item.Submatches
			MsgBox submatch
		Next
		'MsgBox item.Submatches(0)
		'MsgBox item.Submatches(1)
	Next
Else
	MsgBox "No matches found!"
End If
