Option Explicit

Dim fso, a

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("C:\Users\ststoyan\Desktop\Something.txt") Then

a = MsgBox("File already exists. Do you want to overwrite contents?", vbYesNo)

	If a = vbYes Then
	
	fso.CreateTextFile "C:\Users\ststoyan\Desktop\Something.txt"
	
	Else	
	
	WScript.Quit
	
	End If
	
Else

fso.CreateTextFile "C:\Users\ststoyan\Desktop\Something.txt"

End If