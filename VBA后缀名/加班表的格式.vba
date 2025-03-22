Sub ApplyTableFormat()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colIndex As Integer
    
    ' 遍历所有工作表
    For Each ws In ThisWorkbook.Worksheets
        ' 遍历工作表中的所有表格
        For Each tbl In ws.ListObjects
            ' 设置表格的字体
            With tbl.Range.Font
                .Name = "宋体"
                .Size = 12
            End With
            
            ' 设置列宽
            tbl.ListColumns(1).Width = 2 ' 第一列宽2
            tbl.ListColumns(2).Width = 8.38 ' 第二列宽8.38
            tbl.ListColumns(3).Width = 11.5 
            tbl.ListColumns(4).Width = 8.38  
            tbl.ListColumns(5).Width = 20.38 
            tbl.ListColumns(6).Width = 24.63
            tbl.ListColumns(7).Width = 11·
            tbl.ListColumns(8).Width =35.5
            tbl.ListColumns(9).Width = 22
            tbl.ListColumns(10).Width = 12.63
            tbl.ListColumns(12).Width = 5.88
            tbl.ListColumns(13).Width =8.38
           
            ' 设置文本对齐方式为居中
            For colIndex = 1 To tbl.ListColumns.Count
                tbl.ListColumns(colIndex).DataBodyRange.HorizontalAlignment = xlCenter
            Next colIndex
            
            	' 设置行高，这里以设置所有行为例
	' 假设标题行是第一行
	tbl.Range.Rows(1).RowHeight = 51.75 ' 标题行高51.75
	' 假设标题行是第一行
	tbl.Range.Rows(2).RowHeight = 33.75 ' 标题行高51.75

	 ' 找到工作表中的最后一行
    	lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    	' 设置从第一行到最后一行的行高
    	ws.Rows("3:" & lastRow).RowHeight = 30 ' 假设我们设置行高为15


            ' 其他格式设置...
        Next tbl
    Next ws
End Sub