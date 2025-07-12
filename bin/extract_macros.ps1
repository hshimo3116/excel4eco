param(
    [Parameter(Mandatory=$true)]
    [string]$XlsmFile,
    [Parameter(Mandatory=$true)]
    [string]$OutDir
)

if (-not (Test-Path $XlsmFile)) {
    Write-Host "File not found: $XlsmFile"
    exit 1
}

if (-not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($XlsmFile)

try {
    $protection = $workbook.VBProject.Protection
} catch {
    Write-Host "VBProject is inaccessible."
    Write-Host "Check that 'Trust access to the VBA project object model' is enabled. $_"
    $workbook.Close($false)
    $excel.Quit()
    exit 1
}

if ($protection -ne 0) {
    Write-Host "VBA project is protected (Protection=$protection)."
    Write-Host "Remove the project password before exporting."
    $workbook.Close($false)
    $excel.Quit()
    exit 1
}

foreach ($vbComp in $workbook.VBProject.VBComponents) {
    switch ($vbComp.Type) {
        1 { $ext = '.bas' }
        2 { $ext = '.cls' }
        3 { $ext = '.frm' }
        Default { $ext = '.txt' }
    }
    $fileName = $vbComp.Name -replace '[\\/:*?"<>|]', '_'
    try {
        $vbComp.Export((Join-Path $OutDir ($fileName + $ext)))
    } catch {
        Write-Host "Failed to export $($vbComp.Name): $($_.Exception.Message)"
    }
}

$workbook.Close($false)
$excel.Quit()
