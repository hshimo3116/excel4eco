Attribute VB_Name = "Module4"
Option Explicit
Sub ExportAllVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim basePath As String
    Dim fso As Object

    ' ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' このブックのパスを取得
    basePath = ThisWorkbook.Path
    If basePath = "" Then
        MsgBox "このブックは保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If

    ' Modulesフォルダの作成
    exportPath = basePath & "\Modules"
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    ' モジュールのエクスポート
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' 標準モジュール、クラス、フォーム
                vbComp.Export exportPath & "\" & vbComp.Name & ".bas"
        End Select
    Next

    MsgBox "VBAコードのエクスポートが完了しました。" & vbCrLf & "保存先: " & exportPath, vbInformation
End Sub

