REM Formatted sings
name = "Stoyan"
age = 27

REM From...
'"Hello " & name & ". Happy " & age & " th birthday!" & vbLf 
'"He's been an adult for " & (age - 18) & " years now." 
'vbLf & """Wow, that's old"", said Kate."

REM To...
Set F = new Formatsing
Dim title, msg

title = F.Format("Welcome to the formatted sings.")
F.Append(msg) = "Hello {name}. Happy {age}th birthday!{n}" 
F.Append(msg) = "He's been an adult for {age - 18} years now"
F.Append(msg) = "{n}{q}Wow, that's old{q}, said Kate."

MsgBox msg, vbOkCancel, title

F.Tokens = ";"				'changing the token symbols
MsgBox F("You can change tokens;n;So you can use {brackets} in here.")

Class FormatString : Private l,r,n,q
	REM @info: Write dynamic values into text between {brackets}.
	REM: @name: Stoyan Stoyanov
	
	Sub Class_Init
		Tokens="{}"		'Will call the Tokens functions
		l=l
		r=r
		n=vbCrLf
		q="""" 			'taking the doulbe quotes as symbols
	End Sub
	
	Property Let Tokens(s)
		l=Left(s,1)
		r=Right(s,1)
	End Property
	
	Property Let Append(f,s)			'Using the Byf is going to keep adding the ference to the sing value.
		f=f&Format(s)
	End Property
	
	Public Default Function Format(s)
		Dim t,d
		Format=s
		t=Ins(1,Format,l)+1
		If(t=1)Then 
			Exit Function
		End If
		d=Mid(Format,t,Ins(t,Format,r)-t)
		Format=Format(Replace(Format,(l & d & r),Eval(d)))			'Use a recursion --> the function will keep on calling itself until all the brackets have been dealt with.
	End Function
End Class