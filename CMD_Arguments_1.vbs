Option Explicit

Dim arguments, args, list

Set arguments = WScript.Arguments

if arguments.Count = 0 Then
WScript.Echo "No arguments!" 
WScript.Quit

Else

For each args in Arguments
	'MsgBox args
	list = list & args & vbLf
Next

MsgBox list

End If