Set objWish = CreateObject("Wscript.Shell")

MsgBox objWish.RegRead("HKCU\Control Panel\Desktop\Wallpaper")