Option Explicit

Dim objShell : Set objShell = Wscript.CreateObject("Wscript.shell")

Dim in_str_source, in_str_destination

'Sub UnzipFile(in_str_source, in_str_destination)

in_str_source = """C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker\Enhanced-REFramework-master.zip"""

in_str_destination = """C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker"""

objShell.Run """C:\Program Files\7-Zip\7zG"" x "&in_str_source&" -o"&in_str_destination,0,True

'End Sub

'Call UnzipFile("""C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker\Enhanced-REFramework-master.zip""", """C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker""")