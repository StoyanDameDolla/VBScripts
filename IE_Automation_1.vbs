Option Explicit

Dim ie, x

Set ie = CreateObject("InternetExplorer.Application")
Set x = CreateObject("wscript.shell")

ie.Navigate "https://www.google.com"
ie.Toolbar = 0  '0 -- False; 1 - True
ie.StatusBar = 0  '0 -- False; 1 - True
ie.Height = 560
ie.Width = 1000
ie.Top = 0
ie.Left = 0
ie.Resizable = 0
ie.Visible = 1  '0 -- False; 1 - True

Sub WaitForLoad(value)
Do While ie.Busy
wscript.sleep value
Loop
End Sub

Call WaitForLoad(500)

x.sendkeys "cow"
x.sendkeys "{tab}"
x.sendkeys "12345"
x.sendkeys "{enter}"