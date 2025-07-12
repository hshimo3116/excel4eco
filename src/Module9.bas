Attribute VB_Name = "Module9"
Sub Auto_Open()
Attribute Auto_Open.VB_Description = "マクロ記録日 : 2009/2/2  ユーザー名 : 群馬県庁"
Attribute Auto_Open.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Auto_Open Macro
' マクロ記録日 : 2009/2/2  ユーザー名 : 群馬県庁
'

'
    Dim msg As String

    msg = "エコファーマ計画書計算表です。" & vbNewLine & "使いやすくはありません、あしからず"

    MsgBox msg, vbOKOnly + vbInformation, "このシートは・・・"
    
End Sub
