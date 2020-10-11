Option Explicit

Dim fso, objRead, objWrite, cd, content_in_file, o, n, changed_content

cd = Replace(WScript.scriptfullname, WScript.scriptname, "NewFile.txt")

Set fso = CreateObject("Scripting.FileSystemObject")

Set objRead = fso.OpenTextFile(cd, 1)

content_in_file = objRead.ReadAll

o = InputBox(content_in_file, "Replace: Enter a word or phrase you want to replace.")

n = InputBox(content_in_file, "Replace >> " & o & " << With:")

changed_content = Replace(content_in_file, o, n)

objRead.Close

Set objWrite = fso.OpenTextFile(cd, 2)

objWrite.Write changed_content

objWrite.Close

Set objRead = fso.OpenTextFile(cd, 1)

MsgBox objRead.ReadAll, vbInformation, "Here is your new text!"

WScript.Quit