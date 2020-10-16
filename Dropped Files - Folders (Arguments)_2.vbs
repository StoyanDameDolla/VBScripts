Option Explicit
Dim objFso, Args, file

Set objFso = CreateObject("Scripting.FileSystemObject")
Set Args = WScript.Arguments

If Args.Count > 0 Then
	For each file in Args
		If objFso.FileExists(file) Then
			objFso.MoveFile file, "C:\Users\ststoyan\Desktop\Test_Folder\"
		ElseIf objFso.FolderExists(file) Then
			objFso.MoveFolder file, "C:\Users\ststoyan\Desktop\Test_Folder\"
		Else
			MsgBox "Please, drop a file or folder!"
		End If
	Next
Else
MsgBox "Please, drop a file or folder!"
End If