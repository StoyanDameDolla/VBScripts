Option Explicit

Dim files, file, myArray

Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")

Set files = objFso.GetFolder("C:\Users\ststoyan\Desktop\Test_Folder_1\").Files

myArray = Array()

For each file In files

Redim Preserve myArray(UBound(myArray) + 1) 		'Redeclare the array.

myArray(UBound(myArray)) = file.Name

Next

MsgBox myArray(0)
MsgBox Join(myArray, vbLf)