Attribute VB_Name = "Module20Mokuji"

Sub Mokuji()
      Dim i As Integer
   
      '�ڎ��p�ꎞ�V�[�g�ǉ�
      Worksheets.Add before:=Worksheets(1)
   
      '�ڎ��쐬
      Range("A1").value = "---�ڎ�---"
      For i = 2 To Worksheets.Count
            Cells(i, 1).value = Worksheets(i).Name
            ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 1), _
                  Address:="", SubAddress:= _
                  "'" & Worksheets(i).Name & "'!A1"
      Next
   
End Sub


