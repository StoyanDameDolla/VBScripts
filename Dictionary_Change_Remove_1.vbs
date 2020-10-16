Set dict = CreateObject("Scripting.Dictionary")

dict.Add "Steven", "0887815720"
dict.Add "Zlatan", "0887815721"
dict.Add "Ivan", "0887815722"
dict.Add "George", "0887815723"
dict.Add "Peter", "0887815721"
dict.Add "Dragomir", "0887815725"

'allItems = dict.Items()
'WScript.Echo allItems(0)

'dict.Item("Ivan") = "000000000"
'WScript.Echo dict.Item("Ivan")

'MsgBox dict.Keys()(3)
'dict.Key("George") = "Megan"
'MsgBox dict.Keys()(3)

'MsgBox Join(dict.Keys())
'dict.Remove("Peter")
'MsgBox Join(dict.Keys())

MsgBox Join(dict.Items())
dict.Remove("Dragomir")
MsgBox Join(dict.Items())

dict.RemoveAll()
MsgBox Join(dict.Keys())