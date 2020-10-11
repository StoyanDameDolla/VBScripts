Unzip "", ""

Sub Unzip(sSource, sTargetDir)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    if not oFSO.FolderExists(sTargetDir) then oFSO.CreateFolder(sTargetDir)
    Set oShell = CreateObject("Shell.Application")
    Set oSource = oShell.NameSpace(sSource).Items()
    Set oTarget = oShell.NameSpace(sTargetDir)
    oTarget.CopyHere oSource, 16
End Sub