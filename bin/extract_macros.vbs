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
If Not fso.FileExists(xlsmFile) Then
    WScript.Echo "File not found: " & xlsmFile
    WScript.Quit 1
End If
If Not fso.FolderExists(outDir) Then
    fso.CreateFolder outDir
End If

Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False
Set workbook = excelApp.Workbooks.Open(xlsmFile)

' check for VBProject protection or inaccessible project
Dim protection
On Error Resume Next
protection = workbook.VBProject.Protection
If Err.Number <> 0 Then
    WScript.Echo "VBProject is inaccessible. Enable 'Trust access to the VBA project object model'."
    workbook.Close False
    excelApp.Quit
    WScript.Quit 1
End If
On Error GoTo 0
If protection <> 0 Then
    WScript.Echo "VBA project is protected. Unlock the project before exporting."
    workbook.Close False
    excelApp.Quit
    WScript.Quit 1
End If

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
    On Error Resume Next
    vbComp.Export fso.BuildPath(outDir, vbComp.Name & ext)
    If Err.Number <> 0 Then
        WScript.Echo "Failed to export " & vbComp.Name & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
Next

workbook.Close False
excelApp.Quit
