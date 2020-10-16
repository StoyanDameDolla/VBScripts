Option Explicit

Dim objFso, objShell, objApp
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("wscript.shell")
Set objApp = CreateObject("Shell.Application")

'--------Browser for file function---------
Function BrowseForFile()
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
	BrowseForFile = objShell.RegRead(path)
	objShell.RegDelete path
	objFso.DeleteFile temp_folder & "\" & temp_file
End Function

MsgBox BrowseForFile()