Class Clipboard

Private htm, wsh, clip

Private Sub Class_Initialize()
	Set htm = CreateObject("htmlfile")
	Set wsh = CreateObject("WScript.Shell")
End Sub

Public Property Let set_data(strData)
	Set clip = wsh.Exec("clip")
	clip.StdIn.Write strData
	clip.StdIn.Close()
	Do while clip.Status = 0 : WScript.Sleep(200) : Loop
End Property

Public Property Let add_data(strData)
	Set clip = wsh.Exec("clip")
	clip.StdIn.Write get_data
	clip.StdIn.Write strData
	clip.StdIn.Close()
	Do while clip.Status = 0 : WScript.Sleep(200) : Loop 
End Property

Public Property Get get_data()
	get_data = htm.ParentWindow.ClipboardData.GetData("Text")
	If Not TypeName(get_data) = "String" Then
		MsgBox "Invalid clipboard text", vbCritical
		get_data = ""
	End If
End Property

Sub clear_data()
	Set clip = wsh.Exec("clip")
	clip.StdIn.Write ""
	clip.StdIn.Close()
	Do while clip.Status = 0 : WScript.Sleep(200) : Loop
End Sub

Private Sub Class_Terminate()
	Set htm = Nothing
	Set wsh = Nothing
	Set clip = Nothing
End Sub
End Class

Set clip = New Clipboard		'Setting the clip object to inherit the class name. It has access to all function above - as long as they are public!

'MsgBox clip.get_data

clip.set_data = "Something here"

MsgBox clip.get_data

clip.add_data = "as well as something over there!"

MsgBox clip.get_data

clip.clear_data

MsgBox clip.get_data