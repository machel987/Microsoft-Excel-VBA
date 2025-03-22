Sub 拆分单元格()

Dim i, l, m As Integer

'禁止弹出提示的对话框
Application.DisplayAlerts = False

k% = InputBox("请输入拆分单元格所在的列")
'*****************在1995-2006年，excel工作簿包含65536行，但现在的office 2007中工作簿包含1048576行,
'*****************从A列最后一行向上找，找到有数据的行为止
 l = [A65536].End(xlUp).Row      'l记录当前表格的最后一行的行数
 For i = 1 To l
      '判断该单元格是否是合并单元格
       If Cells(i, k).MergeCells = True Then
          m = Cells(i, k).MergeArea.Count                           '记录合并单元格的个数
          Range(Cells(i, k), Cells(i + m - 1, k)).UnMerge      ‘拆分单元格
          Range(Cells(i, k), Cells(i + m - 1, k)).FillDown       ’填充单元格
          i = i + m - 1   
      End If    
 Next
  Application.DisplayAlerts = True
End Sub
