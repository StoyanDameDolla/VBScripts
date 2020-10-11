set x = CreateObject("wscript.shell")

x.run "notepad.exe"
Wscript.sleep 2000
x.sendkeys "Hello, there"
x.sendkeys "{enter}"
x.sendkeys "How are you doing?"
x.sendkeys "%fs"
Wscript.sleep 500
x.sendkeys "SendKeys_2.vbs"
WScript.sleep 500
x.sendkeys "{enter}"