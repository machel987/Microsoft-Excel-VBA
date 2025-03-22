Sub 拆分表格到工作表()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long, i As Long
    Dim key As String
    Dim dict As Object
    Dim Ak As Long

    ' 输入要拆分的列号
    Ak = InputBox("请输入根据哪一列进行拆分，例如输入1, 2, 3, 4")

    ' 设置源工作表
    Set wsSource = ThisWorkbook.Sheets("Sheet1")

    ' 创建字典对象
    Set dict = CreateObject("Scripting.Dictionary")

    ' 获取最后一行的行号
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' 遍历源工作表的每一行
    For i = 2 To lastRow
        ' 获取关键字段值
        key = wsSource.Cells(i, Ak).Value

        ' 如果字典中不存在该关键字段值，则创建新的工作表
        If Not dict.exists(key) Then
            Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsNew.Name = key
            wsSource.Rows(1).Copy Destination:=wsNew.Rows(1)
            dict.Add key, wsNew
        End If

        ' 将当前行复制到对应的工作表中
        wsSource.Rows(i).Copy Destination:=dict(key).Rows(dict(key).Cells(dict(key).Rows.Count, 1).End(xlUp).Row + 1)
    Next i

    ' 提示完成
    MsgBox "拆分完成！"
End Sub

注意：
输入1,代表依据第一列的条件进行拆分到不同的工作表
输入3就是依据第三列的条件进行拆分到不同的工作表
如此类推

为了优化代码，使其更健壮和可靠，以下是逐步的解决方案：

初始化工作簿和工作表：
在代码开始时，检查源工作簿和工作表是否存在。
如果不存在，提示用户并退出。

统一文件名和工作表名称的大小写：
将文件名和工作表名称统一转换为小写或大写，以避免重复或混淆。
处理目标工作表中已有的内容：
在复制工作表范围时，先检查目标工作表中是否有内容。
如果没有内容，再进行复制。

增加错误处理：
在代码开始时，添加错误处理，检查源工作簿和工作表是否存在。
如果不存在，提示用户并退出。

优化代码的可读性和可维护性：
添加注释，解释每一步操作。
使用清晰的变量命名，提高代码的可读性。

修改完善后：
Sub CopyWorksheetRange()
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim dict As Dictionary
    Dim ak As Long

    ' 检查源工作簿是否存在
    If Not ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "Sheet1" Then
        MsgBox "源工作簿未找到，请检查文件路径"
        Exit Sub
    End If

    ' 初始化工作簿和工作表
    Set sourceWorkbook = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set sourceWorksheet = sourceWorkbook.UsedRange

    ' 检查工作表是否存在
    If sourceWorksheet is Nothing Then
        MsgBox "源工作表未找到，请检查工作表位置"
        Exit Sub
    End If

    ' 初始化目标工作簿和工作表
    Set targetWorkbook = ThisWorkbook
    targetWorkbook.Name = "Result"
    Set targetWorksheet = targetWorkbook.Sheets(1)

    ' 统一文件名和工作表名称的大小写
    sourceWorkbook.Name = sourceWorkbook.Name.ToLower()
    sourceWorksheet.Name = sourceWorksheet.Name.ToLower()
    targetWorkbook.Name = targetWorkbook.Name.ToLower()
    targetWorksheet.Name = targetWorksheet.Name.ToLower()

    ' 获取目标工作表的范围
    targetRange = targetWorksheet.UsedRange

    ' 初始化字典
    Set dict = New Dictionary

    ' 遍历源工作表
    For i = 2 To sourceWorksheet.Rows.Count
        key = sourceWorksheet.Cells(i, ak).Value
        If Not dict.ContainsKey(key) Then
            ' 创建新的目标工作表
            Set newWorksheet = targetWorkbook.Sheets.Add
            newWorksheet.Name = key
            ' 复制源工作表的行到目标工作表
            sourceWorksheet.Rows(i).Copy Destination:=newWorksheet.Rows(1)
            ' 移动光标以避免覆盖
            newWorksheet.CurrentCell.Offset(0, 0).Select
            dict.Add key, newWorksheet
        End If
    Next i

    ' 复制使用范围
    If targetRange is Nothing Then
        targetRange = targetWorksheet.UsedRange
    End If

    For i = 2 To sourceWorksheet.Rows.Count
        key = sourceWorksheet.Cells(i, ak).Value
        If Not dict.ContainsKey(key) Then
            ' 复制使用范围到目标工作表
            sourceWorksheet.UsedRange.Copy Destination:=targetRange
            ' 移动光标以避免覆盖
            targetRange.CurrentCell.Offset(0, 0).Select
            dict.Add key, targetRange
        End If
    Next i

    ' 提示复制完成
    MsgBox "复制完成！"
End Sub