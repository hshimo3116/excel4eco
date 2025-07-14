Attribute VB_Name = "modUtilityPack"
Option Explicit

' ==============================
' ���W���[����: modLogger
' �@�\�T�v: Log�V�[�g�Ƀ��O���b�Z�[�W���������ރ��[�e�B���e�B
' ==============================
Public Sub LogWrite(msg As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Log")
    If ws Is Nothing Then
        MsgBox "Log�V�[�g�����݂��܂���", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1).value = Now & " - " & msg
End Sub

' ==============================
' ���W���[����: modConfigReader
' �@�\�T�v: Conf�V�[�g���疼�O��`�iNamed Ranges�j���ꊇ�o�^����
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
' ���W���[����: modMsgBoxEx
' �@�\�T�v: MsgBox�̕\�����e�ƑI�����ʂ����O�ɋL�^����
' ==============================
Public Function MsgBoxLog(prompt As String, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional title As String = "�m�F") As VbMsgBoxResult
    Dim result As VbMsgBoxResult
    result = MsgBox(prompt, buttons, title)
    Call LogWrite("MsgBox [" & title & "]: " & prompt & " �� �I��: " & result)
    MsgBoxLog = result
End Function

' ==============================
' ���W���[����: modRangeHelper
' �@�\�T�v: �w���̍ŏI�s���擾����
' ==============================
Public Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ==============================
' ���W���[����: modEnvironment
' �@�\�T�v: ���[�U�[����Excel�o�[�W�����A�u�b�N��Ԃ̎擾
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
