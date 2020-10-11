Option Explicit

Dim fso, oFile, a

Const Read = 1, Write = 2, Append = 8

Set fso = CreateObject("Scripting.FileSystemObject")

Set oFile = fso.OpenTextFile("C:\Users\ststoyan\Desktop\NewFile.txt", Read, True)

'For a = 1 to 3
'oFile.ReadLine
'Next

'MsgBox oFile.ReadLine

'MsgBox oFile.Read(5)
'MsgBox oFile.ReadLine
'MsgBox oFile.ReadAll
'oFile.Close()

Do until oFile.AtEndOfStream
MsgBox oFile.Read(12)
Loop


