Option Explicit

dim s, t

s = "dd.MM.yyyy"
t = "/"

if inStr(s, "-") Then
  'MsgBox "String contains slash"
	Wscript.Echo Replace(s, "-", t)
else if inStr(s, ".") Then
	Msgbox "No action needed", vbExclamation
	Wscript.Echo Replace(s, ".", t)
else
	Msgbox "No action needed", vbCritical
end if