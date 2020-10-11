Option Explicit

Dim objFso, file_old, file_new, objCreate, objRead, objWrite, line_count, line, content, before, current, after, find, overwrite, ChooseLine, n, i

Const vbRead = 1, vbWrite = 2, vbAppend = 8

Set objFso = CreateObject("Scripting.FileSystemObject")
file_old = "C:\Users\ststoyan\Desktop\File_1.txt"    'File_1.txt is going to be the file we are going to edit.
file_new = Replace(WScript.scriptfullname, WScript.scriptname, "New_File.txt")    'The file instance that will be created. The document is temporarily created in order to write in it.


'------Reading the current file File_1.txt----------
Set objCreate = objFso.OpenTextFile(file_old, 1)
	line_count = 0
Do until objCreate.AtEndOfStream
	line_count = line_count + 1   'incrementing the line count by 1 after each iteration
	line = objCreate.ReadLine
	content = content & line_count & ".)" & line & vbLf
Loop

	ChooseLine = InputBox(content, "Choose a line to edit: 1 - " & line_count)    'line_count represents the total amount of lines from File_1.txt filled with data.
	If isNumeric(ChooseLine) Then
		n = CInt(ChooseLine)
	Else
		MsgBox "WARNING: Value inserted must be numeric!", vbCritical, "Reattempt effort..."
		WScript.Quit
	End If
objCreate.Close

objFso.CreateTextFile file_new

Set objRead = objFso.OpenTextFile(file_old, 1)
For i = 1 to line_count
	if i < n Then
		before = objRead.ReadLine
		Call oldText
	elseif i = n Then
		current = objRead.ReadLine
		Call currentText
	elseif i > n Then
		after = objRead.ReadAll
		Call newText
		Exit For
	End if
Next
objRead.Close


Sub oldText
	Set objWrite = objFso.OpenTextFile(file_new, 8)    'information will be appended, not overwritten!
	objWrite.WriteLine before
	objWrite.Close
End Sub

Sub currentText
	find = InputBox(current, "Replace: please, enter a word/phrase you would like to replace.")
	overwrite = InputBox(current, "Replace - " & find & " - With:")
	Set objWrite = objFso.OpenTextFile(file_new, 8) 	'information will be appended, not overwritten!
	objWrite.WriteLine Replace(current, find, overwrite)
	objWrite.Close
End Sub

Sub newText
	Set objWrite = objFso.OpenTextFile(file_new, 8)		'information will be appended, not overwritten!
	objWrite.Write after
	objWrite.Close
End Sub

