Option Explicit

Dim objShell

Set objShell = CreateObject("wscript.shell")

objShell.AppActivate "Untitled - Notepad"

objShell.SendKeys "(% )(R)"  '-- restore window. 		Alt+Space+R
'objShell.SendKeys "(% )(N)" ' -- minimize window 		Alt+Space+N
objShell.SendKeys "(% )(X)"  ' -- maximize window		Alt+Space+X
objShell.SendKeys "%{F4}"	 ' -- close active window