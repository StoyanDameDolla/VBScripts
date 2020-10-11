Option Explicit

Dim fso : set fso = CreateObject("Scripting.FileSystemObject")

'fso.CopyFile "C:\Users\ststoyan\Desktop\HDigITal_LiveSession_06Okt20.pdf", "C:\Users\ststoyan\Desktop\Test\"
'fso.CopyFolder "C:\Users\ststoyan\Desktop\MergeTablesLogic", "C:\Users\ststoyan\Desktop\Test_2"

'fso.MoveFile "C:\Users\ststoyan\Desktop\Test\HDigITal_LiveSession_06Okt20.pdf", "C:\Users\ststoyan\Desktop\"
'fso.MoveFolder "C:\Users\ststoyan\Desktop\Test_2\MergeTablesLogic", "C:\Users\ststoyan\Desktop\"

'fso.MoveFile "C:\Users\ststoyan\Desktop\Test\*.pdf", "C:\Users\ststoyan\Desktop\"

If fso.FolderExists("C:\Users\ststoyan\Desktop\MergeTablesLogic") Then
fso.MoveFolder "C:\Users\ststoyan\Desktop\MergeTablesLogic", "C:\Users\ststoyan\Desktop\Test_2\"
Else
WScript.Echo "Folder does not exist!"
End If