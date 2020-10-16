Class family

Private new_Name

Property Get lastName
lastName = "Stoyanov"
End Property


Property Get firstName
firstName = new_Name
End Property


Property Let firstName(strFirstName)
new_Name = strFirstName
End Property

Function FullName
	FullName = firstName & " " & lastName
End Function

End Class

Set stoyan = new family
Set dave = new family
Set alex = new family

stoyan.firstName = "Stoyan"

MsgBox stoyan.firstName & " " & stoyan.lastName 
MsgBox stoyan.FullName
'MsgBox dave.firstName = "Dave"
'MsgBox alex.firstName = "Alex"
