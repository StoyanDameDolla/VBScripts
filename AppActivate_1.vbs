Option Explicit

Dim objShell

Set objShell = CreateObject("wscript.shell")

objShell.AppActivate "Untitled - Notepad" 		'Get the title of the application.
WScript.Sleep 500
objShell.SendKeys "Hello, there, my dear friend!"