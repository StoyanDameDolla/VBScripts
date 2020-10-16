Option Explicit

Dim objApp, oldIE, newIE, Window

Set objApp = CreateObject("Shell.Application")
Set oldIE = Nothing

For each Window In objApp.Windows

	If InStr(Window.Name, "Internet Explorer" ) Then

		Set oldIE = Window

	End If
Next

If oldIE is Nothing Then

	Call newWindowInIE

Else

	Call openIE

End If


Sub newWindowInIE
	Set newIE = CreateObject("InternetExplorer.Application")
	newIE.Navigate2 "https://www.youtube.com/", 4096
	newIE.Visible = True
End Sub

Sub openIE
	oldIE.Navigate2 "https://www.youtube.com/", 4096
	oldIE.Visible = True
End Sub