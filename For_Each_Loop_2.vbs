Option Explicit

Dim objFso, objFolder, objFile

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFso.GetFolder("C:\Users\ststoyan\Desktop\Extracted\")

'For each objFile In objFolder.Files 		'write a group or an array'
'MsgBox objFile.Name					
'Next 

For each objFile In objFolder.SubFolders

MsgBox "Folder name: " & objFile.Name & " " & "Folder size: " & objFile.Size

If objFile.Name = "3" Then

Exit For

Else

WScript.Echo objFile.Name

End if

Next

WScript.Echo "Completed!"