考勤表：
' 删除所有的标题
Sub DeleteDescriptions()
    Dim ws As Worksheet
    Dim cell As Range
    Dim descriptions As Variant
    Dim desc As Variant
    
    ' 设置工作表，这里假设是当前活动工作表
    Set ws = ActiveSheet
    
    ' 定义需要删除的描述信息
    descriptions = Array("出勤：○；休息：×（备注栏：病假、事假等文字注明）", _
                        "另： 新入职/辞职请备注", _
                        "一线管理人员：", _
                        "二级部门管理人员：", _
                        "部门负责人：", _
                        "分管领导：", _
                        "清远市清城区顺拓市政园林工程有限公司员工考勤汇总表（2024年11月）")
    
    ' 遍历工作表中的每个单元格
    For Each cell In ws.UsedRange
        ' 检查单元格内容是否包含描述信息
        For Each desc In descriptions
            If InStr(cell.Value, desc) > 0 Then
                ' 删除单元格内容
                cell.Value = ""
            End If
        Next desc
    Next cell
End Sub

除了删除之外，还有一些可以保留部分标题的代码：
Sub DeleteCellsExceptHeaders()
    Dim ws As Worksheet
    Dim cell As Range
    Dim descriptions As Variant
    Dim headers As Variant
    Dim desc As Variant
    Dim header As Variant

    ' 设置工作表，这里假设是当前活动工作表
    Set ws = ActiveSheet
    
    ' 定义需要删除的描述信息
    descriptions = Array("出勤：○；休息：×（备注栏：病假、事假等文字注明）", _
                        "另： 新入职/辞职请备注", _
                        "一线管理人员：", _
                        "二级部门管理人员：", _
                        "部门负责人：", _
                        "分管领导：", _
                        "清远市清城区顺拓市政园林工程有限公司员工考勤汇总表（2024年11月）")
    
    ' 定义需要保留的标题
    headers = Array("序号", "姓名", "部门", "片区", "职务", "班次", "日期")
    
    ' 遍历工作表中的每个单元格
    For Each cell In ws.UsedRange
        ' 检查单元格是否在第一行（标题行）
        If cell.Row = 1 Then
            ' 检查单元格内容是否是标题
            If Not IsInArray(cell.Value, headers) Then
                ' 如果不是标题，则删除单元格
                cell.Delete
            End If
        Else
            ' 检查单元格内容是否包含描述信息
            If IsInArray(cell.Value, descriptions) Then
                ' 删除单元格
                cell.Delete
            End If
        End If
    Next cell
End Sub

Function IsInArray(arr As Variant, value As Variant) As Boolean
    Dim element As Variant
    On Error Resume Next
    IsInArray = Not Err
    For Each element In arr
        If element = value Then
            IsInArray = True
            Exit Function
        End If
    Next element
    On Error GoTo 0
End Function

' 删除除第一行之外的所有的标题
Sub DeleteDescriptionsExceptFirstRow()
    Dim ws As Worksheet
    Dim cell As Range
    Dim descriptions As Variant
    Dim desc As Variant
    Dim i As Long

    ' 设置工作表，这里假设是当前活动工作表
    Set ws = ActiveSheet
    
    ' 定义需要删除的描述信息
    descriptions = Array("出勤：○；休息：×（备注栏：病假、事假等文字注明）", _
                        "另： 新入职/辞职请备注", _
                        "一线管理人员：", _
                        "二级部门管理人员：", _
                        "部门负责人：", _
                        "分管领导：", _
                        "清远市清城区顺拓市政园林工程有限公司员工考勤汇总表（2024年11月）")
    
    ' 从第二行开始遍历工作表中的每个单元格
    For i = 2 To ws.UsedRange.Rows.Count
        For Each cell In ws.Rows(i).Cells
            ' 检查单元格内容是否包含描述信息
            For Each desc In descriptions
                If InStr(cell.Value, desc) > 0 Then
                    ' 删除单元格内容
                    cell.Value = ""
                End If
            Next desc
        Next cell
    Next i
End Sub

加班表：
' 删除所有的标题
Sub DeleteDescriptions()
    Dim ws As Worksheet
    Dim cell As Range
    Dim descriptions As Variant
    Dim desc As Variant
    
    ' 设置工作表，这里假设是当前活动工作表
    Set ws = ActiveSheet
    
    ' 定义需要删除的描述信息
    descriptions = Array("一线管理人员：            二级部门管理人员：            部门负责人：             分管领导：", _
                        "序号", _
                        "日期", _
                        "姓名", _
                        "部门", _
                        "职务", _
                        "加班区域", _
                        "加班事由", _
                        "加班时间", _
                        "小时汇总", _
                        "片区负责人", _
                        "清远市清城区顺拓市政园林工程有限公司加班人员明细表（业务部2024年11月加班明细）")
    
    ' 遍历工作表中的每个单元格
    For Each cell In ws.UsedRange
        ' 检查单元格内容是否包含描述信息
        For Each desc In descriptions
            If InStr(cell.Value, desc) > 0 Then
                ' 删除单元格内容
                cell.Value = ""
            End If
        Next desc
    Next cell
End Sub