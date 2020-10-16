RunAs "wscript.exe", "C:\Users\ststoyan\Desktop\Names_1.vbs",1
RunAs "C:\Users\ststoyan\Desktop\test_1.bat", "cmd.exe",3

Sub RunAs(program, file, show)

	Set objApp = CreateObject("Shell.Application")

	'objApp.ShellExecute "wscript.exe","C:\Users\ststoyan\Desktop\Names_1.vbs",, "runas", 1

	'objApp.ShellExecute "notepad.exe","C:\Users\ststoyan\Desktop\test.txt",, "runas", 1

	objApp.ShellExecute program, file, "runas", show

End Sub