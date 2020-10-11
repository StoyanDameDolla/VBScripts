Set obj = CreateObject("Scripting.FileSystemObject")

Set objFile = obj.GetFile("File_1.txt")

MsgBox objFile.Size