Option explicit

Dim myDate, numDays, nextMonth
nextMonth = desiredMonth + 1 ' desiredMonth is an integer in 1..12
myDate = CDate(nextMonth & "/01/" & desiredYear) ' Could replace with a passed desiredDate instead
myDate = DateAdd("d", -1, myDate)
numDays = Day(myDate)