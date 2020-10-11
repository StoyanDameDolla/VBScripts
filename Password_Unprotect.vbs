Dim objExcel
Set objExcel = WSript.CreateObject("Excel.Application")
objExcel.Workbooks.Open "<<file name and path here>>"
objExcel.ActiveSheet.Unprotect Password:="<<password here>>"
objExcel.Workbooks(1).Save
objExcel.Workbooks.Close
Set objExcel = Nothing