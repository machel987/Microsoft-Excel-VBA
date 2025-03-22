Sub DeleteSampleSections()
    Dim myRange As Range
    Set myRange = ActiveDocument.Content
    
    With myRange.Find
        .ClearFormatting
        .Text = "【示例*】*【适用主题：*】"
        .MatchWildcards = True
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            myRange.Select
            If myRange.Start <> myRange.End Then
                myRange.Delete
            End If
        Loop
    End With
    
    MsgBox "删除完成！共处理" & myRange.Document.Range.ComputeStatistics(wdStatisticCharacters) & "个字符"
End Sub