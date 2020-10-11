Option Explicit

Dim str1, str2

str1 = Month(Date)&"-"&Day(Date)&"-"&Year(Date)
str2 = Month(Date)&"/"&Day(Date)&"/"&Year(Date)

cur_val = InStr(timeStamp,"-")
if cur_val.exist Then
	pass
else
	timeStamp = Month(Date)&"/"&Day(Date)&"/"&Year(Date)
end if

run = MsgBox(timeStamp)
