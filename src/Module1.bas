Attribute VB_Name = "Module1"
Function GetWorkbookPath() As String
    ' �ۑ�����Ă���u�b�N�̃p�X���擾
    If ThisWorkbook.Path <> "" Then
        GetWorkbookPath = ThisWorkbook.Path
    Else
        GetWorkbookPath = "���ۑ�"
    End If
End Function


