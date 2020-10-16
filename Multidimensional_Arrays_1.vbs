Option explicit

Dim array(1,2)

array(0,0) = "Stoyan"
array(0,1) = "Stoyanov"
array(0,2) = 27

array(1,0) = "Ivan"
array(1,1) = "Draganov"
array(1,2) = 55

MsgBox array(0,0) & " " & array(0,1) & " " & array(0,2) & " years of age."
MsgBox array(1,0) & " " & array(1,1) & " " & array(1,2) & " years of age."