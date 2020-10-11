Function QuickZip(path)
  Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")  
  Dim oSap : Set oSap = CreateObject("Shell.Application")
  Dim oWss : Set oWss = CreateObject("WScript.Shell")
  Dim isZip, count, root, base, add, out
  Dim isZipping, isCancelable
  Const NOT_FOUND = 1
  Const NOT_A_ZIP = 2
  Const USER_QUIT = 3
  Select Case oFso.GetExtensionName(path)
    Case "zip"
      If Not oFso.FileExists(path) Then QuickZip = NOT_FOUND : Exit Function Else isZip = True
    Case ""
      If Not oFso.FolderExists(path) Then QuickZip = NOT_FOUND : Exit Function Else isZip = False
    Case Else
      QuickZip = NOT_A_ZIP : Exit Function
  End Select
  If isZip Then 
    root = oFso.GetParentFolderName(path)
    base = oFso.GetBaseName(path)
    out = root & "\" & base
    If oFso.FolderExists(out) Then 
      add = 2 : Do
        out = root & "\" & base & "-" & add
        If Not oFso.FolderExists(out) Then Exit Do
        add = add + 1
      Loop
    End If 
    oFso.CreateFolder(out)   
    oSap.NameSpace(out).CopyHere(oSap.NameSpace(path).Items)
    If FileCount(oSap, path) = FileCount(oSap, out) Then QuickZip = out _
    Else oFso.DeleteFolder out, True : QuickZip = USER_QUIT
  EndIf