Option Explicit
Dim a,b,c

a = 5
b = 3
c = 2

'Msgbox(a - b)
'Msgbox(abs(c - a))
'MsgBox(a/b)
'MsgBox Int(a/b)
'MsgBox Round(a/b)

Randomize
MsgBox Rnd()
MsgBox 100 * Rnd
MsgBox Int(100 * Rnd)
MsgBox Round(100 * Rnd)

If a > b Then
MsgBox "Yes"
Else
MsgBox "No"
End if