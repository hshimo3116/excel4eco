Attribute VB_Name = "Module81ImportAllVBAModules1"
Sub ImportAllVBAModules()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importPath As String
    Dim vbComp As Object
    Dim ext As String
    Dim moduleName As String
    
    ' ファイルシステムオブジェクト作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' srcフォルダのパスを指定
    importPath = ThisWorkbook.Path & "\src"
    If Not fso.FolderExists(importPath) Then
        MsgBox "インポート元フォルダが見つかりません: " & importPath, vbExclamation
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(importPath)
    
    ' 各ファイルを処理
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            moduleName = fso.GetBaseName(file.Name)
            
            ' 既存の同名モジュールがあれば削除
            On Error Resume Next
            Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
            If Not vbComp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
            End If
            On Error GoTo 0
            
            ' インポート処理
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            If Err.Number <> 0 Then
                Debug.Print "インポート失敗: " & file.Name & " (" & Err.Description & ")"
            Else
                Debug.Print "インポート成功: " & file.Name
            End If
            On Error GoTo 0
        End If
    Next

    MsgBox "VBAモジュールのインポートが完了しました。" & vbCrLf & "フォルダ: " & importPath, vbInformation
End Sub

