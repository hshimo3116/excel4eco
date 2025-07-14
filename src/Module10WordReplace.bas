Attribute VB_Name = "Module10WordReplace"
Sub WordKeywordReplace(wdDoc As Object, searchText As String, replaceText As String)
    With wdDoc.Content.Find
        .Text = searchText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=2 ' wdReplaceAll
    End With
End Sub



