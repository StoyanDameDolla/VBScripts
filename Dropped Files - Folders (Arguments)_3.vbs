Option Explicit

Dim Args, arg, List

Set Args = WScript.Arguments

If Args.Count = 0 Then

MsgBox "Please, drop a file or folder!"

Else

For each arg in Args

List = List & arg & vbLf

Next

End If

MsgBox List