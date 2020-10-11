Option Explicit

Dim fso, oFile

Const Read = 1, Write = 2, Append = 8

Set fso = CreateObject("Scripting.FileSystemObject")

Set oFile = fso.OpenTextFile("C:\Users\ststoyan\Desktop\NewFile.txt", Append, True)