Option Explicit

Dim obj 

Set obj = CreateObject("wscript.shell")

MsgBox obj.currentDirectory & "Script_Directory_2.vbs"