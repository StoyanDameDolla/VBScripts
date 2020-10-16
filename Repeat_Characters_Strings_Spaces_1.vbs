Option Explicit

MsgBox "1" & space(20) & "2"
MsgBox "1" & string(50, "|") & "2"
MsgBox "5" & string(5, vbLf) & "8"
MsgBox "5" & StrReverse("Hello, my dear friend!") & "8"
MsgBox MonthName(12,False)