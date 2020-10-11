'Unzip "C:\Users\ststoyan\Downloads\results-export.zip", "C:\Users\ststoyan\HeidelbergCement Group\Robotic Process Automation - Documents\Ariba\TransactionTracker"
Dim Unzip : sSource, sTargetDir
Set Arg = WScript.Arguments
Sub Unzip(sSource, sTargetDir)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    if not oFSO.FolderExists(Arg(1)) then oFSO.CreateFolder(Arg(1))
    Set oShell = CreateObject("Shell.Application")
    Set oSource = oShell.NameSpace(Arg(0)).Items()
    Set oTarget = oShell.NameSpace(Arg(1))
    oTarget.CopyHere oSource, 16
End Sub