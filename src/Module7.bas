Attribute VB_Name = "Module7"
Sub Macro1()
Attribute Macro1.VB_Description = "�}�N���L�^�� : 2008/1/24  ���[�U�[�� : �Q�n����"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro1 Macro
' �}�N���L�^�� : 2008/1/24  ���[�U�[�� : �Q�n����
'

'
    Columns("Q:Q").Select
    Selection.ColumnWidth = 20
End Sub
Sub Macro2()
Attribute Macro2.VB_Description = "�}�N���L�^�� : 2008/1/24  ���[�U�[�� : �Q�n����"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
' �}�N���L�^�� : 2008/1/24  ���[�U�[�� : �Q�n����
'

'
    ActiveWindow.LargeScroll ToRight:=1
    Columns("V:V").Select
    ActiveSheet.Paste
End Sub

Sub �엿���o1()
Attribute �엿���o1.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' �엿���o1 Macro
'

    ' ���������̏�����
    With Worksheets("�엿�R�[�h�\")
        .Range("�엿���o�G���A").AutoFilter Field:=2
        .Range("�엿���o�G���A").AutoFilter Field:=3
        .Range("�엿���o�G���A").AutoFilter Field:=4
        .Range("�엿���o�G���A").AutoFilter Field:=5
    End With

    With Worksheets("�엿�R�[�h�\")
        If .Range("B4").value <> "" Then
            .Range("�엿���o�G���A").AutoFilter Field:=2, Criteria1:="=*" & Range("B4") & "*"
        End If
        If .Range("C4").value <> "" Then
            .Range("�엿���o�G���A").AutoFilter Field:=3, Criteria1:=.Range("C4").Text
        End If
        If .Range("D4").value <> "" Then
            .Range("�엿���o�G���A").AutoFilter Field:=4, Criteria1:=.Range("D4").Text
        End If
        If .Range("E4").value <> "" Then
            .Range("�엿���o�G���A").AutoFilter Field:=5, Criteria1:=.Range("E4").Text
        End If
    End With
End Sub

