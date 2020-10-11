Option Explicit
Dim ie, ipf, wait_value

Set ie = CreateObject("InternetExplorer.Application")

On Error Resume Next

Sub WaitForLoad(wait_value)
Do While IE.Busy
WScript.Sleep wait_value
Loop
End Sub

Sub Fill(x)
Set ipf = ie.Document.All.Item(x)
End Sub

ie.Left = 0
ie.Top = 0
ie.Toolbar = 0
ie.StatusBar = 0
ie.Height = 120
ie.Width = 1020
ie.Resizable = 0

ie.Navigate "https://www.facebook.com/"

wait_value = InputBox("Please, enter a wait time?", 32)
Call WaitForLoad

ie.Visible = True

Call Fill("email")
ipf.Value = "EMAIL GOES HERE"
Call Fill("pass")
ipf.Value = "PASSWORD GOES HERE"
Call Fill("login_form")
ipf.Submit

Call WaitForLoad
ie.Height = 700