'b = (Len(WScript.scriptfullname) - Len(WScript.scriptname))

b = Left(Wscript.scriptfullname, (Len(WScript.scriptfullname) - Len(WScript.scriptname))) ' Obtain the current script path/directory.

MsgBox b