Sub RunAsAdminNoUAC(remove)  
  With CreateObject("WScript.Shell")    
    If WScript.Arguments.length = 0 Then
      If .Run("schtasks /Query /FO CSV /NH /TN """ & WScript.ScriptName & """", 0, True) = 0 Then
        .Run "schtasks /Run /TN """ & WScript.ScriptName & """", 0, True
      Else
        CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & _
        WScript.ScriptFullName & """ AdminArg", "", "runas", 1
      End If
    ElseIf WScript.Arguments.Item(0) = "AdminArg" Then
      .Run "schtasks /Create /SC ONCE /TN """ & WScript.ScriptName & """ /TR ""wscript.exe \""" & _
        WScript.ScriptFullName & "\"" TaskArg"" /ST 00:01 /IT /F /RL HIGHEST", 0, True
      Exit Sub
    ElseIf WScript.Arguments.Item(0) = "TaskArg" Then
      If remove Then .Run "schtasks /Delete /TN """ & WScript.ScriptName & """ /F", 0, True : _
        .Popup "Task Deleted", 2 : WScript.Quit 
      Exit Sub  
    End If
    WScript.Quit
  End With
End Sub

RunAsAdminNoUAC(False)

set objRun = CreateObject("Wscript.Shell")
objRun.Run "cmd /k net user administrator /active:yes"