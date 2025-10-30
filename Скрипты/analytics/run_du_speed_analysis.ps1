Param()

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"
chcp 65001 | Out-Null

$script = Join-Path $PSScriptRoot 'du_speed_analysis.py'
python "$script"

if ($LASTEXITCODE -ne 0) {
    Write-Host "Analysis finished with errors." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Analysis finished successfully." -ForegroundColor Green
}



