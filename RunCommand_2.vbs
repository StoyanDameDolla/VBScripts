Option Explicit

Dim obj, a, b, c

set obj = CreateObject("wscript.shell")

a = MsgBox("Open a program?", vbYesNoCancel+vbQuestion+vbSystemModal)

If a = vbYes Then
	obj.Run "excel.exe"
	b = MsgBox("Open a folder?", vbYesNoCancel+vbQuestion+vbSystemModal)
Else
	b = MsgBox("Open a folder?", vbYesNoCancel+vbQuestion+vbSystemModal)
End If

If b = vbYes Then
	obj.Run "mspaint.exe"
	c = MsgBox("Open a file?", vbYesNoCancel+vbQuestion+vbSystemModal)
Else
	c = MsgBox("Open a file?", vbYesNoCancel+vbQuestion+vbSystemModal)
End If


If c = vbYes Then 
	obj.Run """C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker\Ariba_Parameter_Info.xlsm"""
Else
	WScript.Quit
End If