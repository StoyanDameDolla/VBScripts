Dim multiDimArray(3,2)

multiDimArray(0,0) = "First name:"
multiDimArray(0,1) = "Last name:"
multiDimArray(0,2) = "Age:"

multiDimArray(1,0) = "Stoyan"
multiDimArray(1,1) = "Stoyanov"
multiDimArray(1,2) = "27"

multiDimArray(2,0) = "Zlatan"
multiDimArray(2,1) = "Ivanov"
multiDimArray(2,2) = "25"

multiDimArray(3,0) = "Samuel"
multiDimArray(3,1) = "Lang"
multiDimArray(3,2) = "30"

joined_elements = Split(mJoin(multiDimArray, "|"), "|")
MsgBox joined_elements(4)


'--------Simple join---------
Function mJoin(list, delimiter)
	For each item In list
		allItems = allItems & item & delimiter
	Next
mJoin = allItems
End Function

MsgBox mDJoin(multiDimArray)

'-----------Join for display by applying grid formatting----------
Function mDJoin(list)
For i = 0 to UBound(list) : For m = 0 to UBound(multiDimArray,2)
	If i = 0 and Not m = 0 Then
		allItems = allItems & vbLf
	End If
	allItems = allItems & list(i, m) & vbTab
Next : Next
mDJoin = allItems
End Function
