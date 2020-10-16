Option Explicit

Dim objWMIService

Set objWMIService = GetObject("winmgmts:")

Dim process_list, process, process_run, pass

Set process_list  = objWMIService.ExecQuery("Select * From Win32_Process Where Name = 'explorer.exe'")

For each process In process_list

process.Terminate(1)

Next

Do
pass = InputBox("Please, input a password...")
If pass = "basketball" Then
objWMIService.Get("Win32_Process").Create("explorer.exe")
Exit Do
Else
MsgBox "Please, re-enter password..."
'process.Terminate()
End If
Loop