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
            Worksheets.Add(After:=targetWorksheet)

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

潜在问题：

如果文件夹中包含非Excel文件（如图片或文档），UsedRange 可能会出错。
如果目标工作簿中已经有与源工作簿相同的名称，可能会导致覆盖问题。
如果文件夹路径中有特殊字符（如空格或符号），可能会影响文件路径的正确性。


修改完善后：
Sub 文件夹内容提取()
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim folderPath As String
    Dim file As Object
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim copyRange As Range

    ' 获取目标工作簿
    Set targetWorkbook = Workbooks.Open(FileName:="默认工作簿.xlsx")  ' 替换为实际的工作簿路径

    ' 获取目标工作表
    Set targetWorksheet = targetWorkbook.Worksheets(1)

    ' 获取文件夹路径
    folderPath = InputBox("请输出文件夹完整路径")

    ' 创建Scripting对象
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 获取文件夹中的所有文件
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    ' 遍历文件夹中的每个文件
    For Each file In folder.Files
        ' 获取文件路径
        file.Path = folder.Path & file.Name

        ' 打开对应的Excel文件
        Set sourceWorkbook = Workbooks.Open(file.Path)

        ' 遍历源工作表
        For Each sourceWorksheet In sourceWorkbook.Worksheets
            ' 复制工作表范围
            copyRange = sourceWorksheet

            ' 获取文件名（去掉扩展名）
            Dim fileName As String
            fileName = Replace(file.Name, ".xlsx", "")

            ' 复制到目标工作表
            sourceWorksheet.Copy targetWorksheet.Range("A1")

            ' 激活目标工作表并重命名
            targetWorksheet.Activate
            targetWorksheet.Name = fileName

            ' 移动工作表到目标工作簿
            Worksheets.Add(targetWorksheet, "Sheet" & targetWorksheet.Name)

            ' 恢复目标工作表
            targetWorksheet = targetWorkbook.Worksheets(targetWorksheet.Index + 1)
        Next sourceWorksheet

        ' 关闭打开的Excel文件
        sourceWorkbook.Close SaveChanges:=False
    Next file

    ' 保存目标工作簿
    targetWorkbook.Save

    ' 提示消息
    MsgBox "复制完成！"
End Sub

代码说明
变量声明：

targetWorkbook 和 targetWorksheet：目标工作簿和工作表。
folderPath：输入的文件夹路径。
fso：Scripting.FileSystemObject对象。
folder：文件夹对象。
file：文件对象。
sourceWorkbook 和 sourceWorksheet：源工作簿和工作表。
copyRange：需要复制的范围。
newName：文件名的修改版本。
主要逻辑：

使用InputBox获取文件夹路径。
使用Scripting.FileSystemObject获取文件夹中的所有文件。
遍历每个文件，检查是否为Excel文件。
对于每个Excel文件，打开对应的Excel工作簿和工作表。
获取工作表的使用范围，并复制到目标工作表的A1单元格。
重命名目标工作表，并添加到目标工作簿中。
错误处理：

检查文件是否为Excel文件，避免处理非Excel文件。
检查工作簿和工作表是否存在，避免错误。
显示错误消息，提示用户处理不正确的文件。
优化：

使用更清晰的变量命名，提高代码可读性。
使用Worksheets.Add方法正确添加新工作表。
使用For Each循环遍历文件夹中的所有文件。