Sub 文件夹内容提取()
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim folderPath As String
    Dim file As Object
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim copyRange As Range

    ' 设置目标工作簿和工作表
    Set targetWorkbook = ThisWorkbook
    Set targetWorksheet = targetWorkbook.Worksheets(1)

    ' 获取文件夹路径
    folderPath = InputBox("请输入文件夹完整路径")

    ' 创建FileSystemObject对象
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 获取文件夹对象
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    ' 遍历文件夹中的文件
    Dim shName As String
    For Each file In folder.Files
        ' 打开文件
        Set sourceWorkbook = Workbooks.Open(file.Path)

        ' 遍历文件中的工作表
        For Each sourceWorksheet In sourceWorkbook.Worksheets
            ' 设置复制范围
            Set copyRange = sourceWorksheet.UsedRange

            ' 获取文件名（去除扩展名）
            Dim A As String
            A = Replace(file.Name, ".xlsx", "")

            ' 复制数据到目标工作表
            copyRange.Copy Destination:=targetWorksheet.Range("A1")

            ' 激活目标工作表
            targetWorksheet.Activate

            ' 重命名目标工作表
            targetWorksheet.Name = A

            ' 添加新工作表
            Worksheets.Add After:=targetWorksheet

            ' 激活目标工作表
            targetWorksheet.Activate

            ' 捕获错误
            On Error GoTo infor
            Set targetWorksheet = targetWorkbook.Worksheets(targetWorksheet.Index + 1)
        Next sourceWorksheet

        ' 关闭源工作簿
        sourceWorkbook.Close SaveChanges:=False
    Next file

    ' 保存目标工作簿
    targetWorkbook.Save

    ' 提示完成
    MsgBox "复制完成！"
    Exit Sub

infor:
    MsgBox "请确保目标工作簿是新建的，并正确复制代码后执行。"
End Sub


注意：
输入1,代表依据第一列为条件进行工作簿的拆分，做不同条件的文件分类
输入3就是依据第三列为条件进行工作簿的拆分，做不同条件的文件分类
如此类推

在优化代码的过程中，我需要考虑以下几个方面：
代码的健壮性：确保代码能够处理更多潜在的问题，比如工作表不存在、文件不存在等。
错误处理：在代码开始时添加错误处理，提高代码的健壮性。
代码的可读性和可维护性：添加注释，使用清晰的变量命名，使代码更易读和维护。
性能优化：优化代码的性能，减少不必要的操作。

首先，我会在代码开始时添加错误处理，检查源工作簿和工作表是否存在。如果不存在，提示用户并退出。这将提高代码的健壮性，避免后续操作出错。
其次，我会添加注释，解释每一步操作，使代码更易理解和维护。例如，解释变量的命名和循环的逻辑。
此外，我会优化代码的性能，减少不必要的操作。例如，避免重复的Set语句，使用更高效的变量命名。
最后，我会统一工作簿和工作表的大小写，避免重复或混淆，提高代码的美观性。

在优化过程中，我需要确保代码的功能不变，同时提升其健壮性和美观性。这包括检查所有可能的错误情况，并在必要时进行处理。

总结一下，我的优化步骤如下：
在代码开始时添加错误处理，检查源工作簿和工作表是否存在。
添加注释，解释每一步操作，提高代码的可读性和可维护性。
优化代码的性能，减少不必要的操作。
统一工作簿和工作表的大小写，提高代码的美观性。

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

    ' 初始化字典
    dict = New Dictionary

    ' 复制工作表范围
    For i = 1 To sourceWorksheet.Rows.Count
        key = sourceWorksheet.Cells(i, 1).Value
        If Not dict.ContainsKey(key) Then
            ' 复制工作表范围
            sourceWorksheet.UsedRange.Copy Destination:=targetWorksheet.UsedRange
            ' 移动光标以避免覆盖
            targetWorksheet.CurrentCell.Offset(0, 0).Select
            dict.Add key, targetWorksheet.Name
        End If
    Next i

    ' 提示复制完成
    MsgBox "复制完成！"
End Sub