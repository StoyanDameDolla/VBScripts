Option Explicit

Dim birth

birth = MsgBox("Is today your birthday?", vbYesNo + vbQuestion, "Birthday")

If birth = vbYes Then
	MsgBox "Happy birthday then!", vbInformation
ElseIf birth = vbNo Then 
	MsgBox "Error date!", vbCritical
End If