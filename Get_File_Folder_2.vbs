Option Explicit

Dim objFso, objFile

Set objFso = CreateObject("Scripting.FileSystemObject")

Set objFile = objFso.GetFile("basketball.jpg")

'objFile.Copy "C:\Users\ststoyan\Desktop\Extracted\"
'objFile.Move "C:\Users\ststoyan\Desktop\Extracted\"
'objFile.Delete
set x = objFile.OpenAsTextStream(1)

x.read
x.write
x.writeblanklines
x.writeline

'objFso.CopyFile "source", "destination"'
'objFso.MoveFile "source", "destination"'
'objFso.OpenTextFile "source", IOMode - read, write, append, overwrite'
'objFso.DeleteFile "location"'