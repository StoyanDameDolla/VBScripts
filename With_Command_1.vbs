'set obj = CreateObject("wscript.shell").Run "cmd.exe"

'obj.Run "mspaint.exe"

'With CreateObject("wscript.shell")
'.Run "cmd.exe"	
'.Run "EXCEL.exe"
'.Run "notepad.exe"
'.SendKeys "Hello, there"
'.SendKeys "{enter}"
'.CreateShortcut ("C:\Users\ststoyan\Desktop\Test_Shortcut.lnk")	
'End With

With wscript
.echo "Hello"
.echo "My friend"
.sleep 3000
.quit
End With