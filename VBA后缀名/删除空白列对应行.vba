Sub DeleteRowsWithBlankCells()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 使用当前活动的工作表，也可以指定工作表，如 Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    
    ' 获取最后一行和最后一列
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' 从最后一行开始向上遍历，以便在删除行时不会影响行号
    For i = lastRow To 1 Step -1
        Dim isBlank As Boolean
        isBlank = False ' 假设当前行不是空白行
        
        ' 检查从第一列到最后一列的每个单元格是否为空
        For j = 1 To lastCol
            If IsEmpty(ws.Cells(i, j).Value) Then
                isBlank = True
                Exit For ' 如果找到空白单元格，则跳出循环
            End If
        Next j
        
        ' 如果整行都是空白，则删除该行
        If isBlank Then
            ws.Rows(i).Delete
        End If
    Next i
    
    MsgBox "完成删除空白行操作。"
End Sub