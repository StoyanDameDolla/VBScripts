Option Explicit

Dim obj, oFile, x

Const Read = 1, Write = 2, Append = 8

Set obj = CreateObject("Scripting.FileSystemObject")

Set oFile = obj.OpenTextFile("C:\Users\ststoyan\Desktop\NewFile.txt", Write, True)

x = InputBox("Write to text file")
oFile.Write x 
'oFile.Write "New test file!"
'oFile.Write "Line2"
'oFile.WriteBlankLines(1)
'oFile.Write "Line3"
oFile.Close
Set oFile = obj.OpenTextFile("C:\Users\ststoyan\Desktop\NewFile.txt", Read)
MsgBox oFile.ReadAll
oFile.Close

