Option Explicit

Dim obj

Set obj = CreateObject("wscript.shell")

obj.Run "taskkill /F /IM notepad.exe"