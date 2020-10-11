Option Explicit

Dim objFso, objOtF, current_file, content_in_file

set objFso = CreateObject("Scripting.FileSystemObject")

current_file = Replace(WScript.scriptfullname, WScript.scriptname, "NewFile.txt")

set objOtF = objFso.OpenTextFile(current_file, 1)

content_in_file = objOtF.ReadAll

objOtF.Close

Set objOtF = objFso.OpenTextFile(current_file, 2)

objOtF.Write Replace(content_in_file, "Stoyan", "Steven")

objOtF.Close

WScript.Echo "Operation completed!"