Option Explicit

Dim objShell : Set objShell = CreateObject("WScript.Shell")

Dim fileToZip, fileToExtract 

Function ZippedFile

ZippedFile = "C:\Users\ststoyan\Downloads\Enhanced-REFramework-master.zip"

End Function

Function ExtractFile

ExtractFile = "C:\Users\ststoyan\Desktop\Test_Folder\"

End Function

objShell.Run """C:\Program Files\7-Zip\7z.exe"" x " & ZippedFile & " -o" & ExtractFile, 0, True