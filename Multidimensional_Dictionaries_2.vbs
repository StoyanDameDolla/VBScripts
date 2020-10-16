Set contact = CreateObject("Scripting.Dictionary")

contact.Add "Ben", details("25", "(820)-828 2828")
contact.Add "Jacob", details("56", "(232)-234 9834")


'MsgBox contact.Item("Jacob").item("Age")
MsgBox contact_Details("Ben")
contact.Item("Ben").Item("Age") = 28
MsgBox contact_Details("Ben")


Function details(age, phone)
	Set details = CreateObject("Scripting.Dictionary")
		details.Add "Age", age
		details.Add "Phone", phone
End Function

Function contact_Details(person)
	contact_Details = person & ":" & vbLf
	allKeys = contact.Item(person).Keys
	allItems = contact.Item(person).Items
	
	For i = o To UBound(allItems)
		contact_Details = contact_Details & "    " & allKeys(i) & ":" & " " &   allItems(i) & vbLf
	Next
End Function