Attribute VB_Name = "Module71PrintAll"
Sub PrintAll()
Attribute PrintAll.VB_Description = "マクロ記録日 : 2009/1/30  ユーザー名 : 群馬県庁"
Attribute PrintAll.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Macro3 Macro
' マクロ記録日 : 2009/1/30  ユーザー名 : 群馬県庁
'
' Keyboard Shortcut: Ctrl+Shift+P
'
Dim i
For i = 1 To 23
    Range("C1:D1").Select
    ActiveCell.FormulaR1C1 = i
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Next

    
End Sub
