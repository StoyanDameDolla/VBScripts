Option Explicit

Dim filler, x

set x = CreateObject("wscript.shell")
filler = InputBox("What to search for?")

x.run "chrome.exe"
WScript.sleep 500
x.sendkeys filler
WScript.sleep 500
x.sendkeys "{enter}"