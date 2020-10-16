Set dict = CreateObject("Scripting.Dictionary")

dict.Add "Steven", "0887815720"
dict.Add "Zlatan", "0887815721"
dict.Add "Ivan", "0887815722"
dict.Add "George", "0887815723"
dict.Add "Peter", "0887815721"
dict.Add "Dragomir", "0887815721"

allItems = dict.Items()
allKeys = dict.Keys()

For i = 0 to dict.Count - 1 		'UBound(allItems) or UBound(allKeys)
If allItems(i) = "0887815721" Then
MsgBox allKeys(i)
End If
Next