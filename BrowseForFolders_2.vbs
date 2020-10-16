Option Explicit

Dim objApp

Set objApp = CreateObject("Shell.Application")

Dim objFolder, Path, itemCollection, item


Function SelectFolder(Description)
	Set SelectFolder = objApp.BrowseForFolder(0, Description, 0,0)		'Select folder will be what we obtain as a return value.
	If SelectFolder Is Nothing Then
		WScript.Quit
	End If
End Function

'MsgBox SelectFolder("Select folder for copying:").Self.Path
'MsgBox SelectFolder("Select Folder:").Self.Path

'SelectFolder("Copy the text file here").CopyHere "test.txt"
'Set itemCollection = SelectFolder("Items to move").Items
'SelectFolder("Select target location").MoveHere itemCollection

For each item In SelectFolder("View Items").Items
MsgBox item.Name
Next