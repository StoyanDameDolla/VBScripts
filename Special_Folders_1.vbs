Option Explicit

Dim obj, desktop

Set obj = CreateObject("wscript.shell")

desktop = obj.SpecialFolders("Desktop")

obj.Run desktop & "\BlueJeans"
'obj.Run obj.SpecialFolders("Desktop") & "\BlueJeans"

