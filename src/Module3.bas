Attribute VB_Name = "Module3"
Option Explicit
Function getAct() As String
    Application.Volatile
    getAct = Application.Caller.Parent.Name

End Function

