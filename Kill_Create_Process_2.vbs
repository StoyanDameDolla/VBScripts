Option Explicit

Dim objWMIService 

Set objWMIService = GetObject("winmgmts:")
Dim process_list, process, process_run

Set process_list  = objWMIService.ExecQuery("Select * from Win32_Process")
'("Select * From Win32_Process Where Name = 'notepad.exe'")
'("Select * from Win32_Process")

For each process In process_list

'process.Terminate()
'process_run = process_run & process.Name & vbTab
if process.Name = "notepad.exe" Then

process.Terminate()

End If
Next
MsgBox "terminated"
MsgBox "Rerun of process"
objWMIService.Get("Win32_Process").Create("notepad.exe")