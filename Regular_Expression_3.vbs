str = "Malcom's number is (231)-442 0822. Also Beky has a number I guess and it's 234-837-1928."

Set match = new RegExp

match.Pattern = "\(?\d{3}\)?[\s-]\d{3}[\s-]\d{4}\b|\b\d{10}\b"
match.Global = True

'If match.Test(str) Then
'	MsgBox match.Execute(str).Count
'	MsgBox match.Execute(str).Item(0).FirstIndex			'Returns the elements inside into an array.
'	MsgBox match.Execute(str).Item(0).Length				'Returns the length of the matching pattern.
'	MsgBox match.Execute(str).Item(0).Value					'Returns the whole string matched.
'Else
'	MsgBox "No match found!"
'End If

For i = 0 to match.Execute(str).Count - 1
	MsgBox match.Execute(str).Item(i).Value	
Next

For each item In match.Execute(str)
	MsgBox item.Value	
Next