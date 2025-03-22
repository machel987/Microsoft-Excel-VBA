Sub 从上到下合并单元格()
'
Dim i, l, m As Integer
'禁止弹出提示的对话框
Application.DisplayAlerts = False

k% = InputBox("请输入合并单元格所在的列")

'*****************在1995-2006年，excel工作簿包含65536行，但现在的office 2007中工作簿包含1048576行,
'*****************从A列最后一行向上找，找到有数据的行为止
 l = [A1048576].End(xlUp).Row      'l记录当前表格的最后一行的行数
 
 For i = 1 To l
  
'判断该单元格是否是合并单元格

If Cells(i, k).MergeCells = True Then    
     If Cells(i - m, k) = Cells(i + 1, k) Then     
      m = m + 1                               'M代表一个有多少个相同的单元格   
    Else  
         m = 0
    End If
Else
         m = 0
      If Cells(i, k) = Cells(i + 1, k) Then  
        m = m + 1
      End If
End If 
  Range(Cells(i, k), Cells(i + m, k)).Merge
 Next
 Application.DisplayAlerts = True
 End Sub


Sub 从下到上合并单元格()
Dim i, l As Integer
'禁止弹出提示的对话框
Application.DisplayAlerts = False

k% = InputBox("请输入合并单元格所在的列")
'*****************在1995-2006年，excel工作簿包含65536行，但现在的office 2007中工作簿包含1048576行,
'*****************从A列最后一行向上找，找到有数据的行为止
 l = [A1048576].End(xlUp).Row      'l记录当前表格的最后一行的行数
 
 For i = l To 2 Step -1
     If Cells(i, k) = Cells(i - 1, k) Then
      Range(Cells(i, k), Cells(i - 1, k)).Merge
    End If
 Next
  Application.DisplayAlerts = True  
End Sub



