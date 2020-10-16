Option Explicit

Dim names, name, list

names = Array("John", "Steven", "Dave", "Malcolm", "Peter", "Megan")

For each name In Filter(names, "D")
	list =  list & name & vbLf
Next

MsgBox list