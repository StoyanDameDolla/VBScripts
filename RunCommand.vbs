Option Explicit

Dim obj

Set obj = CreateObject("wscript.shell")

obj.run "mspaint.exe"
obj.run "excel.exe"
obj.run """C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker\"""
