Option Explicit

Dim fso, obj, nDirS, nDirD

Set fso = CreateObject("Scripting.FileSystemObject")

Set obj = CreateObject("WScript.shell")

nDirS = obj.SpecialFolders("MyDocuments")

nDirD = obj.SpecialFolders("Desktop")

fso.CopyFile nDirD & "\HDigITal_LiveSession_06Okt20.pdf", nDirS & "\"

MsgBox "Complete"