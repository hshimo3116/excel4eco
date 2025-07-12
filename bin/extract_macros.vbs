Option Explicit

Dim args, fso, excelApp, workbook, vbComp, outDir, ext, xlsmFile

Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "Usage: cscript extract_macros.vbs <xlsm_file> <output_dir>"
    WScript.Quit 1
End If

xlsmFile = args(0)
outDir = args(1)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(outDir) Then
    fso.CreateFolder outDir
End If

Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
Set workbook = excelApp.Workbooks.Open(xlsmFile)

For Each vbComp In workbook.VBProject.VBComponents
    Select Case vbComp.Type
        Case 1 'vbext_ct_StdModule
            ext = ".bas"
        Case 2 'vbext_ct_ClassModule
            ext = ".cls"
        Case 3 'vbext_ct_MSForm
            ext = ".frm"
        Case Else
            ext = ".txt"
    End Select
    vbComp.Export fso.BuildPath(outDir, vbComp.Name & ext)
Next

workbook.Close False
excelApp.Quit
