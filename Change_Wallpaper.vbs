Set objShell = CreateObject("WScript.Shell")
change_wall = "C:\Users\ststoyan\Desktop\basketball.jpg"

objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", change_wall, "REG_SZ"
objShell.Run "RUNDLL32.exe, user32.dll, UpdatePerUserSystemParameters"