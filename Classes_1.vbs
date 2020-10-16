Class Person
	Private Sub Class_Initialize()
		MsgBox "Starting class."
	End Sub
	
	Private Sub Class_Terminate()
		MsgBox "Ending class"
	End Sub
End class

Set x = New person

x.Class_Initialize
x.Class_Terminate

