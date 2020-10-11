Option Explicit

Dim obj, newLink 

Set obj = CreateObject("WScript.shell")

Set newLink = obj.CreateShortcut("C:\Users\ststoyan\Desktop\AdobeShortcut.lnk")

newLink.TargetPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Acrobat Reader 2017"

newLink.Description = "This is a shortcut of Adobe Acrobat Reader"

newLink.IconLocation = "C:\WINDOWS\System32\SHELL32.dll,27"

newLink.Save
