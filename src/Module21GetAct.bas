Attribute VB_Name = "Module21GetAct"
Function getAct() As String
    Application.Volatile
    getAct = Application.Caller.Parent.Name

End Function

