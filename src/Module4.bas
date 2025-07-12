Attribute VB_Name = "Module4"
Sub ExportAllVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim basePath As String
    Dim fso As Object

    ' �t�@�C���V�X�e���I�u�W�F�N�g�̍쐬
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ���̃u�b�N�̃p�X���擾
    basePath = ThisWorkbook.Path
    If basePath = "" Then
        MsgBox "���̃u�b�N�͕ۑ�����Ă��܂���B��ɕۑ����Ă��������B", vbExclamation
        Exit Sub
    End If

    ' Modules�t�H���_�̍쐬
    exportPath = basePath & "\Modules"
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    ' ���W���[���̃G�N�X�|�[�g
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' �W�����W���[���A�N���X�A�t�H�[��
                vbComp.Export exportPath & "\" & vbComp.Name & ".bas"
        End Select
    Next

    MsgBox "VBA�R�[�h�̃G�N�X�|�[�g���������܂����B" & vbCrLf & "�ۑ���: " & exportPath, vbInformation
End Sub

