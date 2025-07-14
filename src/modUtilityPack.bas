Attribute VB_Name = "modUtilityPack"
Option Explicit

' ==============================
' モジュール名: modLogger
' 機能概要: Logシートにログメッセージを書き込むユーティリティ
' ==============================
Public Sub LogWrite(msg As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Log")
    If ws Is Nothing Then
        MsgBox "Logシートが存在しません", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1).value = Now & " - " & msg
End Sub

' ==============================
' モジュール名: modConfigReader
' 機能概要: Confシートから名前定義（Named Ranges）を一括登録する
' ==============================
Public Sub LoadNamedRangesFromConf()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Conf")
    Dim i As Long
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).value <> "" And ws.Cells(i, 2).value <> "" Then
            ThisWorkbook.Names.Add Name:=ws.Cells(i, 1).value, refersTo:=ws.Cells(i, 2).value
        End If
    Next i
End Sub

' ==============================
' モジュール名: modMsgBoxEx
' 機能概要: MsgBoxの表示内容と選択結果をログに記録する
' ==============================
Public Function MsgBoxLog(prompt As String, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional title As String = "確認") As VbMsgBoxResult
    Dim result As VbMsgBoxResult
    result = MsgBox(prompt, buttons, title)
    Call LogWrite("MsgBox [" & title & "]: " & prompt & " → 選択: " & result)
    MsgBoxLog = result
End Function

' ==============================
' モジュール名: modRangeHelper
' 機能概要: 指定列の最終行を取得する
' ==============================
Public Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ==============================
' モジュール名: modEnvironment
' 機能概要: ユーザー名やExcelバージョン、ブック状態の取得
' ==============================
Public Function GetUserName() As String
    GetUserName = Environ$("Username")
End Function

Public Function GetExcelVersion() As String
    GetExcelVersion = Application.Version
End Function

Public Function IsWorkbookReadOnly() As Boolean
    IsWorkbookReadOnly = ThisWorkbook.ReadOnly
End Function
