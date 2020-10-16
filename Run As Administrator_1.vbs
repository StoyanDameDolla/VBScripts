'Set objApp = CreateObject("Shell.Application")

'objApp.ShellExecute "wscript.exe","C:\Users\ststoyan\Desktop\Names_1.vbs",, "runas", 1

'objApp.ShellExecute "notepad.exe","C:\Users\ststoyan\Desktop\test.txt",, "runas", 1

'objApp.ShellExecute "C:\Users\ststoyan\Desktop\test_1.bat", "cmd.exe",, "runas", 1


RunAsAdmin()

MsgBox "Pause"

Function RunAsAdmin()
	Dim objApp, Args
	Set Args = WScript.Arguments
	If Args.Length = 0 Then
		Set objApp = CreateObject("Shell.Application")
		objApp.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " Run as administrator",,"runas", 1
		WScript.Quit
	End If
End Function