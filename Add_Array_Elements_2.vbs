Option Explicit

Dim animals_array, objShell

animals_array = Array("Bear", "Wolf", "Snail")

animals_array = AddToArray(animals_array)

With CreateObject("wscript.shell")
.Run "notepad.exe"
WScript.sleep 1000
.SendKeys Join(animals_array, vbLf)
End With

Function AddToArray(current_array)
	Dim value
	if IsArray(current_array) Then
		Do
		value = InputBox(Join(current_array, vbLf), "Add element to array.")
		ReDim Preserve current_array(UBound(current_array) + 1)
		current_array(UBound(current_array)) = value
		Loop until value = ""
	End If
	AddToArray = current_array
End Function

