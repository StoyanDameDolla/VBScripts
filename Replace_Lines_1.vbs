Option Explicit

Dim fso, cd, cd2, objCreate, objRead, objWrite, line_count, before, current, after, change, i

Const vbvRead = 1, vbWrite = 2, vbAppend = 8

Set fso = CreateObject("Scripting.FileSystemObject")
cd = "C:\Users\ststoyan\Desktop\File_1.txt"
cd2 = Replace(WScript.scriptfullname, Wscript.scriptname, "File_2.txt")

fso.CreateTextFile cd2

Set objCreate = fso.OpenTextFile(cd, 1)
line_count = 0
Do until objCreate.AtEndOfStream
line_count = line_count + 1
objCreate.SkipLine
Loop
objCreate.Close

Set objRead = fso.OpenTextFile(cd, 1)
For i = 1 to line_count
	if i < 2 Then
		before = objRead.ReadLine
		Call oldText
	elseif i = 2 Then
		current = objRead.ReadLine
		Call currentText
	elseif i > 2 Then
		after = objRead.ReadAll
		Call newText
		Exit For
	End If
Next

objRead.Close

fso.CopyFile cd2, cd, True 'True means it will overwrite contents
fso.DeleteFile cd2

Sub oldText
	Set objWrite = fso.OpenTextFile(cd2, 8)
	objWrite.WriteLine before
	objWrite.Close
End Sub

Sub currentText
	change = Replace(current, "text file", "document file")
	Set objWrite = fso.OpenTextFile(cd2, 8)
	objWrite.WriteLine change
	objWrite.Close
End Sub

Sub newText
	Set objWrite = fso.OpenTextFile(cd2, 8)
	objWrite.Write after
	objWrite.Close
End Sub