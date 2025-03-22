1：
Sub 单元合并()
    Dim i As Integer
    Dim startRow As Integer
    Dim endRow As Integer
    Dim col As Integer

    ' 从第30行开始，每两行合并一次
    startRow = 30
    endRow = 51 ' 假设您需要处理到第51行

    ' 循环处理每两行
    For i = startRow To endRow Step 2
        ' 循环处理A到E列
        For col = 1 To 6
            ' 合并当前行和下一行的单元格
            Range(Cells(i, col), Cells(i + 1, col)).Merge
            ' 设置合并后的单元格属性
            With Range(Cells(i, col), Cells(i + 1, col))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
        Next col
    Next i
End Sub

2：
Sub 单元合并()
    Dim i As Integer
    Dim startRow As Integer
    Dim endRow As Integer
    Dim col As Integer

    ' 从第13行开始，每两行合并一次
    startRow = 13
    endRow = 51 ' 假设您需要处理到第51行

    ' 循环处理每两行
    For i = startRow To endRow Step 2
        ' 循环处理A到E列
        For col = 1 To 6
            ' 合并当前行和下一行的单元格
            Range(Cells(i, col), Cells(i + 1, col)).Merge
            ' 设置合并后的单元格属性
            With Range(Cells(i, col), Cells(i + 1, col))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
        Next col
    Next i
End Sub