Sub Resize_Columns_And_Rows_No_Header2()

Dim currentSheet As Worksheet

    Set currentSheet = ActiveSheet

    Dim sheet As Worksheet
    For Each sheet In ActiveWorkbook.Worksheets
        With sheet
            With .Cells.Rows
                .WrapText = True
                .VerticalAlignment = xlCenter
                .EntireRow.AutoFit
            End With '.Cells.Rows
            .Columns.EntireColumn.AutoFit
        End With 'sheet
    Next sheet

    currentSheet.Activate

End Sub

Sub WrapAlighText()

Range("A8:U8").WrapText = True
Range("A8:U8").HorizontalAlignment = xlCenter
Range("A8:U8").VerticalAlignment = xlCenter

End Sub

Sub SetCustomColumnWidth()

Worksheet("Orderlist").Range("N8:Q8").ColumnWidth = 12.43

End Sub
