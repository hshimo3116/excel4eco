Attribute VB_Name = "Module1"
Option Explicit
Function GetWorkbookPath() As String
    ' 保存されているブックのパスを取得
    If ThisWorkbook.Path <> "" Then
        GetWorkbookPath = ThisWorkbook.Path
    Else
        GetWorkbookPath = "未保存"
    End If
End Function


