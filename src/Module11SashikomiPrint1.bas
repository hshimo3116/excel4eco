Attribute VB_Name = "Module11SashikomiPrint1"
Sub Wordへ一括差し込み処理()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim templatePath As String
    Dim outputPath As String
    Dim i As Long
    Dim lastRow As Long
    Dim tag As String
    Dim value As String

    ' Confシートから設定値を取得（2行目以降）
    With Sheets("Conf")
        templatePath = .Range("B2").value
        outputPath = .Range("B4").value
    End With

    ' テンプレートファイルの存在確認
    If Dir(templatePath) = "" Then
        MsgBox "テンプレートファイルが見つかりません：" & vbCrLf & templatePath, vbExclamation
        Exit Sub
    End If

    ' 出力フォルダの整形と確認
    If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"
    If Dir(outputPath, vbDirectory) = "" Then
        MsgBox "出力フォルダが見つかりません：" & vbCrLf & outputPath, vbExclamation
        Exit Sub
    End If

    ' Word起動
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    ' テンプレートを開く
    Set wdDoc = wdApp.Documents.Open(templatePath)

    ' タグ一覧を Tags シートから読み取る（B列とD列、3行目以降）
    With Sheets("Tags")
        lastRow = .Cells(.Rows.Count, 2).End(xlUp).Row ' B列の最終行
        For i = 3 To lastRow ' ← 3行目から処理
            tag = .Cells(i, 2).value    ' B列（2列目）
            value = .Cells(i, 4).value  ' D列（4列目）
            If tag <> "" Then
                Call Word置換(wdDoc, tag, value)
            End If
        Next i
    End With

    ' 保存ファイル名を生成
    Dim saveFilePath As String
    Dim fileName As String
    fileName = Range("NUMBER").value & Range("申請者名").value & "別記様式1_" & Format(Now, "yyyymmdd_hhmmss") & ".docx"
    saveFilePath = outputPath & fileName

    ' 保存
    wdDoc.SaveAs2 fileName:=saveFilePath
    MsgBox "出力が完了しました：" & vbCrLf & saveFilePath

    ' （任意）Word終了処理
    ' wdDoc.Close
    ' wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

