Option Explicit

Dim fso, drives, drive

Set fso = CreateObject("Scripting.FileSystemObject")
Set drives = WScript.Arguments

For each drive In drives

MsgBox fso.GetFolder(drive & ":\").Subfolders.Count, vbOkOnly, "Folders in " & drive

Next