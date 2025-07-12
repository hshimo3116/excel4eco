Attribute VB_Name = "Module1"
Sub Auto_Open()
'
' Auto_Open Macro
' �}�N���L�^�� : 2009/2/2  ���[�U�[�� : shimo-hi
' �e�X�g�����R�����g����� test comment

'
    Dim msg As String

    msg = "�G�R�t�@�[�}�v�揑�v�Z�\�ł��B" & vbNewLine & "�g���₷���͂���܂���A�������炸"

    MsgBox msg, vbOKOnly + vbInformation, "���̃V�[�g�́E�E�E"
    Call ���O��`_Config�V�[�g����ꊇ�o�^
    
End Sub

Function GetWorkbookPath() As String
    ' �ۑ�����Ă���u�b�N�̃p�X���擾
    If ThisWorkbook.Path <> "" Then
        GetWorkbookPath = ThisWorkbook.Path
    Else
        GetWorkbookPath = "���ۑ�"
    End If
End Function


Sub ���O��`_Config�V�[�g����ꊇ�o�^()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim nm As String, val As String, refersTo As String
    Dim cellRef As String
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("config")
    
    i = 2 ' �� �w�b�_�[�s�i1�s�ځj���X�L�b�v
    
    Do While ws.Cells(i, 1).value <> ""
        nm = Trim(ws.Cells(i, 1).value)
        val = Trim(ws.Cells(i, 2).Formula)     ' B��F�l
        refersTo = Trim(ws.Cells(i, 3).Formula) ' C��F�Q�Ɣ͈�
        
        ' ���O���󔒂܂��͖����Ȃ�X�L�b�v
        If nm <> "" Then
            ' ����������΍폜�i�㏑���Ή��j
            On Error Resume Next
            wb.Names(nm).Delete
            On Error GoTo 0
            
            ' ���O��`�FB��D��A����C��
            If val <> "" Then
                wb.Names.Add Name:=nm, refersTo:=val
            ElseIf refersTo <> "" Then
                wb.Names.Add Name:=nm, refersTo:=refersTo
            End If
            
            ' D��F��`���ꂽ���O�̎Q�Ɛ�i������\���j
            On Error Resume Next
            If wb.Names(nm).RefersToRange Is Nothing Then
                ' �萔�ȂǂŎQ�Ƃ��Ȃ��ꍇ
                cellRef = wb.Names(nm).refersTo
            Else
                cellRef = wb.Names(nm).RefersToRange.Address(External:=True)
            End If
            On Error GoTo 0
            ws.Cells(i, 4).value = cellRef
        End If
        
        i = i + 1
    Loop
    
    MsgBox "���O��`���������܂����B", vbInformation
End Sub

