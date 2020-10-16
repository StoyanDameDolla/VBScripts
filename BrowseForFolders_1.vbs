Option Explicit

Dim objApp

Set objApp = CreateObject("Shell.Application")

Dim objFolder, Path


Function SelectFolder
	Set SelectFolder = objApp.BrowseForFolder(0, "Select Folder:", 0,0)		'Select folder will be what we obtain as a return value.

	'If objFolder Is Nothing Then
	If SelectFolder Is Nothing Then

	'MsgBox "Canceled"

	WScript.Quit

	'Else

	'MsgBox objFolder.Title
	'MsgBox objFolder.Self.Path
	'Path = objFolder.Self.Path & "\"
	'MsgBox Path
	End If
End Function

MsgBox SelectFolder.Self.Path