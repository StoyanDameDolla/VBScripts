'x = "Stoyan Stoyanov"
'MsgBox TypeName(x)
'x = 30
'MsgBox TypeName(x)
'x = 2.4542
'MsgBox TypeName(x)
'x = Null
'MsgBox TypeName(x)
'x = Empty
'MsgBox TypeName(x)
'x = False
'MsgBox TypeName(x)


set x = CreateObject("WScript.shell")
MsgBox Not IsObject(x)

y = 12
MsgBox IsNumeric(y)

If VarType(null) = vbEmpty Then
MsgBox "Yes"
Else
MsgBox "No"
End if

x = InputBox("Please, enter a value", "Input box for value", "default value goes here...")
If IsNumeric(x) Then
	If VarType(x) = vbInteger Then
		MsgBox "Yes"
	Else
		MsgBox "No"
	End if
Else
	MsgBox "Not numeric!"
End If