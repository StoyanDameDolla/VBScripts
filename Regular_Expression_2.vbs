str = "Is John Huckmann in town! John is looking for him! No more John related searches!"

Set name = New RegExp

name.Pattern = "\bJohn\b"
'name.IgnoreCase = False
name.Global = True

MsgBox name.Replace(str, "Steven")
'str = name.Replace(str, "Adam")

'MsgBox str

If name.Test(str) = "Jo" Then
	MsgBox "String exists!"
Else
	MsgBox "String does not exist!"
End If
