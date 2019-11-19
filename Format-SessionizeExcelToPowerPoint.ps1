param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$excelFileFullPath,
    [Parameter(Mandatory=$True)]
    [string]$powerPointFileFullPath
)
.\Format-SessionizeExcelToPowerPoint.vbs (Join-Path $pwd $excelFileFullPath) (Join-Path $pwd $powerPointFileFullPath)