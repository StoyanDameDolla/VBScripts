Set dict = CreateObject("Scripting.Dictionary")

dict.Add "Steven", "0887815720"
dict.Add "Zlatan", "0887815721"
dict.Add "Ivan", "0887815722"
dict.Add "George", "0887815723"

'MsgBox dict.Item("Ivan")

allItems = dict.Items() 'Collection of dictionary items, entered into an array!
allKeys = dict.Keys()
'MsgBox allItems(3)
'MsgBox allKeys(0)
'MsgBox allKeys(1)
'MsgBox allKeys(2)
'MsgBox allKeys(3)
'MsgBox Join(allItems)
'MsgBox Join(allKeys)
MsgBox dict.Exists("Hrisi")
MsgBox dict.Count