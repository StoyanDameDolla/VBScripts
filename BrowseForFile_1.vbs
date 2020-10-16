Option Explicit

'------Declare all variables-------

Dim objFso, objShell, objApp
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objApp = CreateObject("Shell.Application")

'--------Browse for folder---------

Function SelectFolder()
	Dim objFolder
	Set objFolder = objApp.BrowseForFolder(0,"Select folder:",0,0)
	If objFolder Is Nothing Then
	MsgBox "Canceled"
	WScript.Quit
	Else
	SelectFolder = objFolder.Self.Path & "\"
	End If
End Function

'--------Browse for file-----------

Function SelectFile()
	Dim temp_folder, temp_file, path
	Set temp_folder = objFso.GetSpecialFolder(2)  'GetSpecialFolder --> will create a temporary folder. A random path can be created.
	temp_file = objFso.GetTempName() & ".hta" '.hta --> stands for Hyper Text Application
	path = "HKCU\Volatile Environment\MsgResp"
	With temp_folder.CreateTextFile(temp_file)
		.Write "<input type=file name=f>" & _
		"<script>f.click();(new ActiveXObject('wscript.shell'))" & _
		".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
		"close();</script>"
		.Close
	End With
	objShell.Run temp_folder & "\" & temp_file, 0, True	'0 --> window of application will not be shown; 1 --> app window will be shown; True value is set --> will wait for previous processes to finish before executing the current one.
	If objShell.RegRead(path) = "" Then
		objShell.RegDelete path
		objFso.DeleteFile temp_folder & "\" & temp_file
		WScript.Quit
	End If
	SelectFile = objShell.RegRead(path)
	objShell.RegDelete path
	objFso.DeleteFile temp_folder & "\" & temp_file

End Function

'-------Move/Copy file from one directory into another------
objFso.MoveFile SelectFile, SelectFolder