'Dim Args, arg, DroppedFiles, file

Set Args = WScript.Arguments

If Args.Count > 0 Then

	DroppedFiles = Array()

	For each file In Args

		ReDim Preserve DroppedFiles(UBound(DroppedFiles) + 1)
		
		DroppedFiles(UBound(DroppedFiles)) = file

	Next
Else
	MsgBox "Please, drop a file or folder!"
End If

MsgBox Join(DroppedFiles, vbLf)
MsgBox DroppedFiles(1)