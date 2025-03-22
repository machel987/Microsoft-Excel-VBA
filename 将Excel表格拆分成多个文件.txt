Sub SplitWorkbook()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim LastRow As Long
    Dim FileName As String
    Dim Path As String

    '获取当前工作簿
    Set wb = ActiveWorkbook

    '获取工作簿中第一个工作表
    Set ws = wb.Worksheets(1)

    '获取工作表中最后一行的行号
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    '获取工作簿的保存路径
    Path = wb.Path

    '循环遍历工作表中的每一行
    For i = 2 To LastRow
        '获取当前行的第一个单元格的值
        FileName = ws.Cells(i, 1).Value

        '如果当前行第一个单元格的值为空，则跳过
        If FileName = "" Then
            GoTo ContinueLoop
        End If

        '新建一个工作簿
        Workbooks.Add.SaveAs Path & "\" & FileName & ".xlsx"

        '将当前行的数据复制到新工作簿中
        ws.Range(ws.Cells(i, 1), ws.Cells(i, ws.Columns.Count)).Copy
        Windows(FileName & ".xlsx").Activate
        ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteValues

        '关闭新工作簿
        ActiveWorkbook.Close savechanges:=False

ContinueLoop:
    Next i

    '提示用户拆分完成
    MsgBox "拆分完成"
End Sub