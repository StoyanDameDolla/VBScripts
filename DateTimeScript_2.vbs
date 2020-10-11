Option Explicit

Wscript.Echo "Today is " & myDate(now)

' date formatted as your request
Function myDate(dt)
    dim d,m,y, sep
    sep = "-"
    ' right(..) here works as rpad(x,2,"0")
    d = right("0" & datePart("d",dt),2)
    m = right("0" & datePart("m",dt),2)
    y = datePart("yyyy",dt)
    myDate= m & sep & d & sep & y
End Function