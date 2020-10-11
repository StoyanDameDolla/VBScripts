Option Explicit

Dim objFso, oTextFile 

Const Read = 1, Write = 2, Append = 8

set objFso = CreateObject("Scripting.FileSystemObject")

set oTextFile = objFso.OpenTextFile("C:\Users\ststoyan\Desktop\NewFile.txt", Append)

oTextFile.Write " " & "I am from Bourgas, Bulgaria."
oTextFile.WriteBlankLines(3)
oTextFile.WriteLine "My wife is called Hristina. She is 27 years old."
'oTextFile.WriteBlankLines()
