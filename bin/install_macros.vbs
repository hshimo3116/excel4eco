Option Explicit

Dim args, fso, excelApp, workbook, vbComp, srcDir, file, ext, xlsmFile

Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "Usage: cscript install_macros.vbs <xlsm_file> <source_dir>"
    WScript.Quit 1
End If

xlsmFile = args(0)
srcDir = args(1)

Set fso = CreateObject("Scripting.FileSystemObject")

Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
Set workbook = excelApp.Workbooks.Open(xlsmFile)

' remove existing modules (except document modules)
For Each vbComp In workbook.VBProject.VBComponents
    If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then
        workbook.VBProject.VBComponents.Remove vbComp
    End If
Next

If fso.FolderExists(srcDir) Then
    For Each file In fso.GetFolder(srcDir).Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            workbook.VBProject.VBComponents.Import file.Path
        End If
    Next
End If

workbook.Save
workbook.Close
excelApp.Quit
