Option Explicit

Dim objFso, objFile

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.GetFile("C:\Users\ststoyan\Desktop\NewFile.txt")

if objFile.Attributes And 2 Then    ' add 16 to check if it a directory
objFile.Attributes = objFile.Attributes - 2
Else
objFile.Attributes = objFile.Attributes + 2
End if

WScript.Echo "Finished"