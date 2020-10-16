Option Explicit

Dim a, b, c

a = 55464.2342
b = 18.8546
c = .86

'WScript.Echo FormatNumber(a, 0)
'WScript.Echo FormatCurrency(b, 2)
'WScript.Echo FormatPercent(c, 0)

MsgBox FormatDateTime("20-12-2020", vbLongDate)
MsgBox FormatDateTime("May, 5, 2020", vbGeneralDate)
MsgBox FormatDateTime(Time(), vbShortTime)