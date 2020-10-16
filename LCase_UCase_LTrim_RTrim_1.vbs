Option Explicit

Dim var, var1

var = "This is a test message. Cantaloupe..."
var1 = "     Hello, there    "

'MsgBox LTrim(var)
'MsgBox RTrim(var)

If InStr(LCase(var), "cantaloupe") Then
	MsgBox "Yes"
Else
	MsgBox "No"
End If

MsgBox LTrim(var1) & "boy"
MsgBox RTrim(var1) & "boy"
MsgBox Trim(var1) & "boy"