$taskRoot = Split-Path -Parent $PSScriptRoot
$taskInput = (Get-ChildItem -LiteralPath $taskRoot -Filter '*.docx' | Select-Object -First 1).FullName
$taskPdf = Join-Path $PSScriptRoot 'render\alltime.pdf'
$taskWord = $null
$taskDoc = $null
try {
    $taskWord = New-Object -ComObject Word.Application
    $taskWord.Visible = $false
    $taskWord.DisplayAlerts = 0
    $taskDoc = $taskWord.Documents.Open($taskInput, $false, $true)
    $taskDoc.ExportAsFixedFormat($taskPdf, 17)
    Write-Output ('PAGES=' + $taskDoc.ComputeStatistics(2))
} finally {
    if ($null -ne $taskDoc) { $taskDoc.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskDoc) }
    if ($null -ne $taskWord) { $taskWord.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskWord) }
}
