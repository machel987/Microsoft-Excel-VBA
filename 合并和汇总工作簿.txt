'合并工作簿
'合并是把文件夹的表格文件放同一个表文件里变成多个工作簿
Option Explicit
 方法一：
Sub 合并工作簿()
    '定义文件变量
    Dim filestoopen, ft
    Dim wk As Workbook
    '文件数量变量
    Dim x As Integer
    '关闭屏幕刷新
    Application.ScreenUpdating = False
    '选择需要合并的文件
    filestoopen = Application.GetOpenFilename(filefilter:="MIc(*.xlsx), *.xlsx", MultiSelect:=True, Title:="请选择需要合并的文件")
    '未选定文件时进行提示
    If TypeName(filestoopen) = "boolean" Then
        MsgBox "未选定文件。"
    End If
    '初始化
    x = 1
 
    '逐一打开选定的工作簿，将每个工作簿的第一个工作表复制到当前工作簿中
    While x <= UBound(filestoopen)
        Set wk = Workbooks.Open(Filename:=filestoopen(x))
        wk.Sheets(1).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wk.Close SaveChanges:=False
        Set wk = Nothing
        x = x + 1
    Wend
 
    '打开屏幕刷新
    Application.ScreenUpdating = True
    MsgBox "合并完成"
 
End Sub

加上表格的报数修改后：
Sub 合并工作簿()
    '定义文件变量
    Dim filestoopen, ft
    Dim wk As Workbook
    '文件数量变量
    Dim x As Integer
    '工作簿名称累积字符串
    Dim workbookNames As String
    '关闭屏幕刷新
    Application.ScreenUpdating = False
    '选择需要合并的文件
    filestoopen = Application.GetOpenFilename(filefilter:="Excel Files (*.xlsx), *.xlsx", MultiSelect:=True, Title:="请选择需要合并的文件")
    '未选定文件时进行提示
    If TypeName(filestoopen) = "boolean" Then
        MsgBox "未选定文件。"
        Exit Sub
    End If
    '初始化
    x = 1
    workbookNames = ""
 
    '逐一打开选定的工作簿，将每个工作簿的第一个工作表复制到当前工作簿中
    While x <= UBound(filestoopen)
        Set wk = Workbooks.Open(Filename:=filestoopen(x))
        workbookNames = workbookNames & wk.Name & Chr(13) ' 添加工作簿名称并换行
        wk.Sheets(1).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wk.Close SaveChanges:=False
        Set wk = Nothing
        x = x + 1
    Wend
 
    '打开屏幕刷新
    Application.ScreenUpdating = True
    '显示合并的工作簿数量和名称
    MsgBox "共合并了" & x - 1 & "个工作簿。如下：" & Chr(13) & workbookNames, vbInformation, "提示"
End Sub
 

 方法二：
Sub 工作薄间工作表合并()
  
Dim FileOpen
Dim X As Integer
Application.ScreenUpdating = False
FileOpen = Application.GetOpenFilename(FileFilter:="Microsoft Excel文件(*.xls*),*.xls", MultiSelect:=True, Title:="合并工作薄")
X = 1
While X <= UBound(FileOpen)
Workbooks.Open Filename:=FileOpen(X)
Sheets().Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
X = X + 1
Wend
ExitHandler:
Application.ScreenUpdating = True
Exit Sub
  
errhadler:
MsgBox Err.Description
End Sub

加上表格的报数修改后：
Sub 工作薄间工作表合并()
    Dim FileOpen As Variant
    Dim X As Integer
    Dim workbookNames As String
    Application.ScreenUpdating = False
    FileOpen = Application.GetOpenFilename(FileFilter:="Microsoft Excel文件(*.xls*),*.xls*", MultiSelect:=True, Title:="合并工作薄")
    If TypeName(FileOpen) = "Boolean" Then
        MsgBox "未选定文件。"
        Exit Sub
    End If
    X = 1
    workbookNames = ""
    While X <= UBound(FileOpen)
        Workbooks.Open Filename:=FileOpen(X)
        workbookNames = workbookNames & ActiveWorkbook.Name & Chr(13) ' 添加工作簿名称并换行
        Sheets().Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveWorkbook.Close SaveChanges:=False
        X = X + 1
    Wend
    Application.ScreenUpdating = True
    '显示合并的工作簿数量和名称
    MsgBox "共合并了" & X - 1 & "个工作簿。如下：" & Chr(13) & workbookNames, vbInformation, "提示"
End Sub

'汇总工作表
Option Explicit
 
Sub 汇总工作表()
    Dim j As Integer
    Application.ScreenUpdating = False
 
    For j = 1 To Sheets.Count
        If Sheets(j).Name <> ActiveSheet.Name Then
            Sheets(j).UsedRange.Copy Cells(Range("A65536").End(xlUp).Row + 1, 1)
        End If
    Next
 
    Range("A1").Select
    '打开屏幕刷新
    Application.ScreenUpdating = True
End Sub

加上表格的报数修改后：
'汇总工作表
Option Explicit

Sub 汇总工作表()
    Dim j As Integer
    Dim lastRow As Long
    Dim totalSheets As Integer ' 用于记录汇总的工作表数量
    Application.ScreenUpdating = False

    ' 找到活动工作表中的最后一行
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

    totalSheets = 0 ' 初始化汇总的工作表数量

    ' 从第一个工作表开始遍历到工作簿中的最后一个工作表
    For j = 1 To Sheets.Count
        If Sheets(j).Name <> ActiveSheet.Name Then
            ' 复制非活动工作表的已用范围到活动工作表的下一行
            Sheets(j).UsedRange.Copy Destination:=ActiveSheet.Cells(lastRow + 1, 1)
            ' 更新活动工作表的最后一行
            lastRow = lastRow + Sheets(j).UsedRange.Rows.Count
            totalSheets = totalSheets + 1 ' 汇总一个工作表，计数器加1
        End If
    Next

    ' 如果汇总了工作表，则显示汇总结果
    If totalSheets > 0 Then
        MsgBox "共汇总了 " & totalSheets & " 个工作表。", vbInformation, "汇总完成"
    Else
        MsgBox "没有其他工作表可以汇总。", vbInformation, "汇总完成"
    End If

    ' 打开屏幕刷新
    Application.ScreenUpdating = True
End Sub
