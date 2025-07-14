Attribute VB_Name = "Module81ImportAllVBAModules1"
Sub ImportAllVBAModules()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importPath As String
    Dim vbComp As Object
    Dim ext As String
    Dim moduleName As String
    
    ' �t�@�C���V�X�e���I�u�W�F�N�g�쐬
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' src�t�H���_�̃p�X���w��
    importPath = ThisWorkbook.Path & "\src"
    If Not fso.FolderExists(importPath) Then
        MsgBox "�C���|�[�g���t�H���_��������܂���: " & importPath, vbExclamation
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(importPath)
    
    ' �e�t�@�C��������
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            moduleName = fso.GetBaseName(file.Name)
            
            ' �����̓������W���[��������΍폜
            On Error Resume Next
            Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
            If Not vbComp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
            End If
            On Error GoTo 0
            
            ' �C���|�[�g����
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            If Err.Number <> 0 Then
                Debug.Print "�C���|�[�g���s: " & file.Name & " (" & Err.Description & ")"
            Else
                Debug.Print "�C���|�[�g����: " & file.Name
            End If
            On Error GoTo 0
        End If
    Next

    MsgBox "VBA���W���[���̃C���|�[�g���������܂����B" & vbCrLf & "�t�H���_: " & importPath, vbInformation
End Sub

