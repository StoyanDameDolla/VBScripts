a = MsgBox("Pandora Box",vbAbortRetryIgnore + vbExclamation + vbDefaultButton3 + vbSystemModal,"Gadget:")

if a = vbAbort Then MsgBox "Quit", vbCritical
if a = vbRetryCancel Then
End If