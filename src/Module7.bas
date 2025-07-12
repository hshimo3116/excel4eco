Attribute VB_Name = "Module7"
Sub Macro1()
Attribute Macro1.VB_Description = "マクロ記録日 : 2008/1/24  ユーザー名 : 群馬県庁"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro1 Macro
' マクロ記録日 : 2008/1/24  ユーザー名 : 群馬県庁
'

'
    Columns("Q:Q").Select
    Selection.ColumnWidth = 20
End Sub
Sub Macro2()
Attribute Macro2.VB_Description = "マクロ記録日 : 2008/1/24  ユーザー名 : 群馬県庁"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
' マクロ記録日 : 2008/1/24  ユーザー名 : 群馬県庁
'

'
    ActiveWindow.LargeScroll ToRight:=1
    Columns("V:V").Select
    ActiveSheet.Paste
End Sub

Sub 肥料抽出1()
Attribute 肥料抽出1.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' 肥料抽出1 Macro
'

    ' 検索条件の初期化
    With Worksheets("肥料コード表")
        .Range("肥料抽出エリア").AutoFilter Field:=2
        .Range("肥料抽出エリア").AutoFilter Field:=3
        .Range("肥料抽出エリア").AutoFilter Field:=4
        .Range("肥料抽出エリア").AutoFilter Field:=5
    End With

    With Worksheets("肥料コード表")
        If .Range("B4").value <> "" Then
            .Range("肥料抽出エリア").AutoFilter Field:=2, Criteria1:="=*" & Range("B4") & "*"
        End If
        If .Range("C4").value <> "" Then
            .Range("肥料抽出エリア").AutoFilter Field:=3, Criteria1:=.Range("C4").Text
        End If
        If .Range("D4").value <> "" Then
            .Range("肥料抽出エリア").AutoFilter Field:=4, Criteria1:=.Range("D4").Text
        End If
        If .Range("E4").value <> "" Then
            .Range("肥料抽出エリア").AutoFilter Field:=5, Criteria1:=.Range("E4").Text
        End If
    End With
End Sub

