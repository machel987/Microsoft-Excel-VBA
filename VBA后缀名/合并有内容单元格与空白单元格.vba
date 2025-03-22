Sub 从下到上合并有内容单元格与空白单元格()
    Dim i, l As Integer
    Application.DisplayAlerts = False

    k% = InputBox("请输入需要操作的列号") ' 获取用户输入的列号
    l = [A1048576].End(xlUp).Row ' 获取最后一行的行号

    ' 从第1行开始向下遍历
    For i = 1 To l Step 2 ' 每次跳2行，因为有内容单元格和空白单元格是成对出现的
        ' 检查当前单元格是否有内容，且下一个单元格是否为空
        If Cells(i, k).Value <> "" And Cells(i + 1, k).Value = "" Then
            ' 合并当前单元格和下一个空白单元格
            Range(Cells(i, k), Cells(i + 1, k)).Merge
        End If
    Next

    Application.DisplayAlerts = True
End Sub


Sub 从上到下合并单元格合并有内容单元格与空白单元格()
    Dim i, l As Integer
    Application.DisplayAlerts = False

    k% = InputBox("请输入需要操作的列号") ' 获取用户输入的列号
    l = [A1048576].End(xlUp).Row ' 获取最后一行的行号

    ' 从第1行开始向下遍历
    For i = 1 To l
        ' 检查当前单元格是否有内容，且下一个单元格是否为空
        If Cells(i, k).Value <> "" And Cells(i + 1, k).Value = "" Then
            ' 合并当前单元格和下一个空白单元格
            Range(Cells(i, k), Cells(i + 1, k)).Merge
        End If
    Next

    Application.DisplayAlerts = True
End Sub

注意：输入的结果为数字
数字相对应所在的列数

比如：数字1相对应列A