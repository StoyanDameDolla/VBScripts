Option Explicit

Dim Animals(2), names, i

animals(0) = "horse"
animals(1) = "fox"
animals(2) = "rabbit"

'MsgBox animals(1)
names = Array("John", "Steven", "Dave", "Malcolm") 

'For each i In names
'	MsgBox(i)
'Next
'MsgBox names(0)
'MsgBox names(1)
'MsgBox names(2)

MsgBox Join(animals, vbLf)
MsgBox Join(names, vbTab)

For i = LBound(names) To UBound(names)		'LBound can be set to 0;
	MsgBox names(i)
Next