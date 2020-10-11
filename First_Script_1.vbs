Option Explicit

Dim fso, objWriteToFile, objReadToFile, file_to_save, current_data, create_data

Const Read = 1, Write = 2, Append = 8

Set fso = CreateObject("Scripting.FileSystemObject")

file_to_save = Replace(WScript.scriptfullname, WScript.scriptname, "%username%_File.txt")

'--------------------------------------------------------------------------------------'

Sub Refresh

Set objReadToFile = fso.OpenTextFile(file_to_save, Read)    'True == 1; False == 0

current_data = objReadToFile.ReadAll

objReadToFile.Close

End Sub

'-------Main-------'
'Information to be inserted into Input box'
'------------------'

Sub Main

Do

create_data = InputBox(current_data, "Information input = [*] to quit")

If create_data = "*" Then

WScript.Quit

ElseIf create_data = "" Then

	Set objWriteToFile = fso.OpenTextFile(file_to_save, Append, True)

	objWriteToFile.WriteBlankLines(1)

	objWriteToFile.Close

	Call Refresh

Else 

	Set objWriteToFile = fso.OpenTextFile(file_to_save, Append, True)

	objWriteToFile.Write create_data

	objWriteToFile.Close

	Call Refresh

End If

Loop

End Sub

'----------WRITE INTO FILE-----------'

If not fso.FileExists(file_to_save) Then

	Call Main

Else 

If fso.GetFile(file_to_save).Size = 0 Then

	Call Main

Else

	Call Refresh

	Call Main

End If

End If

WScript.Quit