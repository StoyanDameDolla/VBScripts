Option Explicit

dim dateValue, myStringDateCorrect, myStringDateWrong1, myStringDateWrong2
	
	myStringDateWrong1 = "."
	myStringDateWrong2 = "-"
	myStringDateCorrect = "/"
Do
dateValue = InputBox("Please, enter a date format", "Date Window Input", "Enter date here...")

if InStr (1,Date) = myStringDateWrong1 Then
	Wscript.Echo Replace(dateValue, myStringDateWrong1, myStringDateCorrect)
	MsgBox dateValue
	exit do
elseif InStr (1,Date) = myStringDateWrong2 Then
	Wscript.Echo Replace(dateValue, myStringDateWrong2, myStringDateCorrect)
	MsgBox dateValue
	exit do
elseif InStr (1,Date) = myStringDateCorrect Then
	MsgBox dateValue
	exit do
	'Msgbox ("Date format is suitable"), vbExclamation
else
	MsgBox ("Please, enter a valid date format!"), vbExclamation
end if
MsgBox dateValue
Loop until InStr(1,dateValue) = myStringDateCorrect

