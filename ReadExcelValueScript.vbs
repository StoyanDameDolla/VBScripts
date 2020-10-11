Option Explicit

Dim objExcel,ObjWorkbook,objsheet, readCell

Set objExcel = CreateObject("Excel.Application")
Set ObjWorkbook = objExcel.Workbooks.Open("C:\Users\ststoyan\Documents\UiPath\DateFormatting.xlsx") 'path to spreadsheet can be parameterized
Set objsheet = objExcel.ActiveWorkbook.Worksheets(1)

readCell = objsheet.Cells(2,2).value

WScript.StdOut.Write(readCell)  

Set objSheet = Nothing
ObjWorkbook.Close
Set ObjWorkbook = Nothing
objExcel.Quit 
Set objExcel = Nothing