Dim objEnv : Set objEnv = CreateObject("WScript.Shell")

'temp = objEnv.Environment("User").Item("TEMP")
'MsgBox temp
'MsgBox objEnv.ExpandEnvironmentStrings(temp)
'MsgBox objEnv.ExpandEnvironmentStrings("%username%")
'MsgBox objEnv.ExpandEnvironmentStrings("%programfiles%")
'MsgBox objEnv.ExpandEnvironmentStrings("%temp%")

Set user = objEnv.Environment("User")
Set system = objEnv.Environment("System")
'user.Item("TEST_ITEM") = "STOYAN"
'MsgBox "Pause"
'user.Remove("TEST_ITEM")
MsgBox system.Item("Path")
MsgBox Split(system.Item("Path"), ";")(2)
System.Item("Path") = System.Item("Path") & ";C:\myLocation\path"