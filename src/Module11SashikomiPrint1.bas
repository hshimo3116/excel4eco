Attribute VB_Name = "Module11SashikomiPrint1"
Sub Word�ֈꊇ�������ݏ���()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim templatePath As String
    Dim outputPath As String
    Dim i As Long
    Dim lastRow As Long
    Dim tag As String
    Dim value As String

    ' Conf�V�[�g����ݒ�l���擾�i2�s�ڈȍ~�j
    With Sheets("Conf")
        templatePath = .Range("B2").value
        outputPath = .Range("B4").value
    End With

    ' �e���v���[�g�t�@�C���̑��݊m�F
    If Dir(templatePath) = "" Then
        MsgBox "�e���v���[�g�t�@�C����������܂���F" & vbCrLf & templatePath, vbExclamation
        Exit Sub
    End If

    ' �o�̓t�H���_�̐��`�Ɗm�F
    If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"
    If Dir(outputPath, vbDirectory) = "" Then
        MsgBox "�o�̓t�H���_��������܂���F" & vbCrLf & outputPath, vbExclamation
        Exit Sub
    End If

    ' Word�N��
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    ' �e���v���[�g���J��
    Set wdDoc = wdApp.Documents.Open(templatePath)

    ' �^�O�ꗗ�� Tags �V�[�g����ǂݎ��iB���D��A3�s�ڈȍ~�j
    With Sheets("Tags")
        lastRow = .Cells(.Rows.Count, 2).End(xlUp).Row ' B��̍ŏI�s
        For i = 3 To lastRow ' �� 3�s�ڂ��珈��
            tag = .Cells(i, 2).value    ' B��i2��ځj
            value = .Cells(i, 4).value  ' D��i4��ځj
            If tag <> "" Then
                Call Word�u��(wdDoc, tag, value)
            End If
        Next i
    End With

    ' �ۑ��t�@�C�����𐶐�
    Dim saveFilePath As String
    Dim fileName As String
    fileName = Range("NUMBER").value & Range("�\���Җ�").value & "�ʋL�l��1_" & Format(Now, "yyyymmdd_hhmmss") & ".docx"
    saveFilePath = outputPath & fileName

    ' �ۑ�
    wdDoc.SaveAs2 fileName:=saveFilePath
    MsgBox "�o�͂��������܂����F" & vbCrLf & saveFilePath

    ' �i�C�ӁjWord�I������
    ' wdDoc.Close
    ' wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

