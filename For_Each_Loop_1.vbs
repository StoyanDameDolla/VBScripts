Option Explicit

Dim loop_counter

For loop_counter = 0 to 4

if loop_counter = 2 Then
exit For
End if
MsgBox "Looping"
Next

MsgBox "Successfully finished!"