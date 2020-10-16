Set list = CreateObject("System.Collections.ArrayList")

list.add "Steven" 
list.add "Stoyan"
list.add "Ivan"

'For Each item In Split("ben alica mark john")
'	list.add(item) 
'Next

Sub addRange(list, range)
	For each item In range
		list.add item
	Next
End Sub

addRange list, Split("Ben Alex Mark John")

'MsgBox list(4)
'MsgBox list.Item(4)
'MsgBox list.Count
list.Sort
MsgBox Join(list.ToArray)
list.Reverse
MsgBox Join(list.ToArray)
MsgBox list.Contains("Alex")
MsgBox list.IndexOf("Alex", 0)
list.Insert 0, "Garry"
MsgBox Join(list.ToArray)
list.Remove "Steven"
MsgBox Join(list.ToArray)

Set names = list.Clone

list.Clear
MsgBox Join(list.ToArray)
MsgBox Join(names.ToArray)