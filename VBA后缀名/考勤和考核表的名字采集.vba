步骤：
先“合并工作簿”，再复制这个名字的采集，最后“按A列数据修改表名称”
同理：考勤表的合并和采集也是一样的道理

'合并工作簿
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

Sub 按A列数据修改表名称AAAA()

    On Error Resume Next '忽略错误继续执行VBA代码，避免出现错误消息
 
    Application.Calculation = xlCalculationAutomatic '手动重算
 
    Dim i%
 
    For i = 1 To Sheets.Count
 
    Sheets(i).Name = Cells(i, 1).Text
 
    Next
 
    On Error GoTo 0 '恢复正常的错误提示
 
    Application.Calculation = xlCalculationAutomatic '自动重算
 
End Sub

一线生产人员汇总（231）
一线管理人员考核表（19）
一线生产管理人员考核表（25）
一线生产人员（仓管2）
洲心-黄尚谦17
洲心西-汤家丽18
洲心东-李红娥23
横荷-麦国良29
龙塘-麦国良13
东城-邓健伟21
奥体-何维21
凤城-徐立明11
黑臭水体-杨永坚18
车队-冯智明26
工程队-朱金胜4
机修-周党能1
修剪队-刘海英9
修树一队-赖广清4
修树二队-王大安4
门卫-朱杰3
苗圃-黄威7

考勤表的名字采集：
业务部（道路）12月考勤汇总表279
道路部-李傲霜24
洲心市区-黄尚谦21
洲心西-汤家丽21
洲心东片区-李红娥25
东城-邓健伟25
东城奥体-何维24
凤城-徐立明13
横荷-麦国良33
横荷龙塘片区-麦国良14
黑臭水体-杨永坚18
车队-冯智明28
工程队-朱金胜4
机动维修队-周党能2
修剪打草队-刘海英9
修树一队-赖广清4
修树二队-王大安4
保安队-朱杰3
苗圃场-黄威7