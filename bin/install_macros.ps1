param(
    [Parameter(Mandatory=$true)]
    [string]$XlsmFile,
    [Parameter(Mandatory=$true)]
    [string]$SrcDir
)

if (-not (Test-Path $XlsmFile)) {
    Write-Host "File not found: $XlsmFile"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($XlsmFile)

foreach ($vbComp in @($workbook.VBProject.VBComponents)) {
    if ($vbComp.Type -eq 1 -or $vbComp.Type -eq 2 -or $vbComp.Type -eq 3) {
        $workbook.VBProject.VBComponents.Remove($vbComp)
    }
}

if (Test-Path $SrcDir) {
    Get-ChildItem -Path $SrcDir | ForEach-Object {
        if ($_.Extension -in '.bas','.cls','.frm') {
            $workbook.VBProject.VBComponents.Import($_.FullName)
        }
    }
}

$workbook.Save()
$workbook.Close()
$excel.Quit()
