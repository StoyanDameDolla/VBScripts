Set objWsh = CreateObject("WScript.Shell")

RunAsAdmin()

objWsh.RegWrite "HKEY_CLASSES_ROOT\VBSFile\Shell\Edit\Icon", "C:\WINDOWS\System32\SHELL32.dll,141", "REG_SZ"

Function RunAsAdmin()
 Dim objAPP
  If WScript.Arguments.length = 0 Then
  Set objAPP = CreateObject("Shell.Application")
  objAPP.ShellExecute "wscript.exe", """" & _
  WScript.ScriptFullName & """" & " RunAsAdministrator",,"runas", 1
  WScript.Quit
  End If
End Function
