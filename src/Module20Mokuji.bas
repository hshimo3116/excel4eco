Attribute VB_Name = "Module20Mokuji"

Sub Mokuji()
      Dim i As Integer
   
      '目次用一時シート追加
      Worksheets.Add before:=Worksheets(1)
   
      '目次作成
      Range("A1").value = "---目次---"
      For i = 2 To Worksheets.Count
            Cells(i, 1).value = Worksheets(i).Name
            ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 1), _
                  Address:="", SubAddress:= _
                  "'" & Worksheets(i).Name & "'!A1"
      Next
   
End Sub


