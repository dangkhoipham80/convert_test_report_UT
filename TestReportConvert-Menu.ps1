#Requires -Version 5.1
# Opens the JavaScript web UI (local server). Same folder as serve_ui.py.
# ASCII only for Windows PowerShell 5.1 file encoding.
param()

$ErrorActionPreference = "Stop"
$BaseDir = $PSScriptRoot
$Server = Join-Path $BaseDir "serve_ui.py"

if (-not (Test-Path -LiteralPath $Server)) {
    Write-Host "Not found: $Server" -ForegroundColor Red
    return
}

function Get-Python {
    if (Get-Command python -ErrorAction SilentlyContinue) { return "python" }
    if (Get-Command py -ErrorAction SilentlyContinue) { return "py" }
    return $null
}

$py = Get-Python
if (-not $py) {
    Write-Host "Python not in PATH. Install from https://www.python.org/ (Add to PATH)." -ForegroundColor Red
    return
}

if (-not (Test-Path -LiteralPath (Join-Path $BaseDir "web\index.html"))) {
    Write-Host "Missing web\index.html" -ForegroundColor Red
    return
}

Set-Location -LiteralPath $BaseDir
Write-Host "Starting web UI. Browser will open. Press Ctrl+C here to stop the server." -ForegroundColor Cyan
Write-Host "URL: http://127.0.0.1:8765/" -ForegroundColor DarkGray
& $py $Server
