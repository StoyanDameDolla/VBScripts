Option Explicit

Const vbRead = 1, vbWrite = 2, vbAppend = 8
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim readVBSfile : Set readVBSfile = objFso.OpenTextFile("Execute - Read from another vbscript file_1.vbs",1)
Dim readVBS_2file : Set readVBS_2file = objFso.OpenTextFile("Names_1.vbs",1)

Execute readVBSfile.ReadAll()
Execute readVBS_2file.ReadAll()

readVBSfile.Close
readVBS_2file.Close

MsgBox AddVars(2,3)
MsgBox SubstractVars(7,-9)
MsgBox MultiplyVars(9, 8)

MsgBox firstName
MsgBox lastName