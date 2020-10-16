Option Explicit

Dim message, i, array, n

message = "Hello, Johh! How are you? Well, I am feeling awesome! Terrific!"

'For i = 1 to Len(message)
'	MsgBox Mid(message,i,5)
'	MsgBx Split(str)
'Next

'array = Split(message, "!")

'For i = 0 to UBound(array)
'	MsgBox array(i)
'Next

'MsgBox InStr(message,"John")
If InStr(15, message, "Johh") Then
MsgBox "Yes"
Else
MsgBox "No"
End If

'MsgBox array(0)
'MsgBox array(1)
'MsgBox array(2)
'MsgBox array(3)
'MsgBox array(4)
'MsgBox array(5)