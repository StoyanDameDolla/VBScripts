Option Explicit

MsgBox Date()
MsgBox Time()
MsgBox Now()

'MsgBox DateAdd("d", + 3, Date)
'MsgBox DateAdd("m", + 7, Date)
'MsgBox DateAdd("yyyy", -2, Date)

'MsgBox DateDiff("s", DateAdd("n", -10, Now), Now)
'MsgBox DateDiff("ww", "01/09/2020", Now)

MsgBox DatePart("w", Now)
MsgBox DatePart("q", "01/07/2020")

'MsgBox Year(Now)
'MsgBox Month(Now)
'MsgBox Weekday(Now)
'MsgBox Day(Now)
'MsgBox Hour(Now)
'MsgBox Minute(Now)
'MsgBox Second(Now)

'MsgBox MonthName(10, True)
'MsgBox MonthName(10, False)

'MsgBox WeekdayName(6, True)
'MsgBox WeekdayName(7, False)