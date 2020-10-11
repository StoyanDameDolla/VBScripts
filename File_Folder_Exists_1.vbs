Option Explicit

Dim fso : set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("C:\Users\ststoyan\Desktop\HDigITal_LiveSession_06Okt20.pdf") Then
Wscript.echo "Yes"
Else
Wscript.echo "No"
End If