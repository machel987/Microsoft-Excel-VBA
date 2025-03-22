Sub 考勤表的数据填充():
    Dim i, j
    For i = 4 To 451
        For j = 7 To 37
            If Cells(i, j) = "×" Then
                GoTo NextIteration
            End If
            Cells(i, j) = "○"
NextIteration:
        Next j
    Next i
End Sub

Sub 考勤表的数据填充2()
     Rem 只用于单一的表格
    Dim ws As Worksheet
    ' 请确保工作表名称正确
    Set ws = ThisWorkbook.Sheets("道路部-李傲霜24") 
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long

    ' 确定数据范围
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    ' 循环遍历每一行
    For i = 2 To lastRow ' 假设第一行是标题，从第二行开始
        ' 循环遍历每一天的考勤状态
        For j = 7 To 36 ' 假设从第7列开始是考勤状态
            ' 如果单元格为空，则填充"○"
            If IsEmpty(ws.Cells(i, j).Value) Then
                ws.Cells(i, j).Value = "○"
            End If
        Next j
    Next i
End Sub

Sub 考勤表的数据填充3()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long

    sheetNames = Array("道路部-李傲霜24", "洲心市区-黄尚谦20", "洲心西-汤家丽21", "洲心东片区-李红娥25", "东城-邓健伟23", "东城奥体-何维24", "凤城-徐立明12", "横荷-麦国良33", "横荷龙塘片区-麦国良11", "黑臭水体-杨永坚16", "车队-冯智明28", "工程队-朱金胜4", "机动维修队-周党能2", "修剪打草队-刘海英9", "修树一队-赖广清4", "修树二队-王大安3", "保安队-朱杰3", "苗重场-黄威10")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

        For i = 2 To lastRow
            For j = 7 To 36
                If IsEmpty(ws.Cells(i, j).Value) Then
                    ws.Cells(i, j).Value = "○"
                End If
            Next j
        Next i
    Next sheetName
End Sub