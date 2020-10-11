Option Explicit 

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

fso.DeleteFile "C:\Users\ststoyan\Desktop\Something.txt"