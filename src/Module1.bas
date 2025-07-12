Attribute VB_Name = "Module1"
Sub Auto_Open()
'
' Auto_Open Macro
' マクロ記録日 : 2009/2/2  ユーザー名 : shimo-hi
' テスト試験コメントを入力 test comment

'
    Dim msg As String

    msg = "エコファーマ計画書計算表です。" & vbNewLine & "使いやすくはありません、あしからず"

    MsgBox msg, vbOKOnly + vbInformation, "このシートは・・・"
    Call 名前定義_Configシートから一括登録
    
End Sub

Function GetWorkbookPath() As String
    ' 保存されているブックのパスを取得
    If ThisWorkbook.Path <> "" Then
        GetWorkbookPath = ThisWorkbook.Path
    Else
        GetWorkbookPath = "未保存"
    End If
End Function


Sub 名前定義_Configシートから一括登録()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim nm As String, val As String, refersTo As String
    Dim cellRef As String
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("config")
    
    i = 2 ' ← ヘッダー行（1行目）をスキップ
    
    Do While ws.Cells(i, 1).value <> ""
        nm = Trim(ws.Cells(i, 1).value)
        val = Trim(ws.Cells(i, 2).Formula)     ' B列：値
        refersTo = Trim(ws.Cells(i, 3).Formula) ' C列：参照範囲
        
        ' 名前が空白または無効ならスキップ
        If nm <> "" Then
            ' 同名があれば削除（上書き対応）
            On Error Resume Next
            wb.Names(nm).Delete
            On Error GoTo 0
            
            ' 名前定義：B列優先、次にC列
            If val <> "" Then
                wb.Names.Add Name:=nm, refersTo:=val
            ElseIf refersTo <> "" Then
                wb.Names.Add Name:=nm, refersTo:=refersTo
            End If
            
            ' D列：定義された名前の参照先（文字列表示）
            On Error Resume Next
            If wb.Names(nm).RefersToRange Is Nothing Then
                ' 定数などで参照がない場合
                cellRef = wb.Names(nm).refersTo
            Else
                cellRef = wb.Names(nm).RefersToRange.Address(External:=True)
            End If
            On Error GoTo 0
            ws.Cells(i, 4).value = cellRef
        End If
        
        i = i + 1
    Loop
    
    MsgBox "名前定義が完了しました。", vbInformation
End Sub

