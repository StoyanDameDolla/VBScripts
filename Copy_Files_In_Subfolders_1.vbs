Option Explicit

Dim objFSO, objStart, objType, objEnd, objFile, folder, sub_folder
Set objFSO = CreateObject("Scripting.FileSystemObject")

objStart = "C:\Users\ststoyan\Desktop\Test_Folder_1"
objEnd = "C:\Users\ststoyan\Desktop\Destination_Folder\"

For each objFile In objFSO.GetFolder(objStart).Files
	objFile.Copy objEnd
Next

Call SearchSubfolders(objFSO.GetFolder(objStart))

Function SearchSubfolders(folder)
	For each sub_folder In folder.SubFolders
		For each objFile In objFSO.GetFolder(sub_folder.Path).Files
			objFile.Copy objEnd
		Next
	Call SearchSubfolders(sub_folder)
	Next
End Function

MsgBox "Process finished", vbInformation