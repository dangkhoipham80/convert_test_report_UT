# Cài lệnh vào profile PowerShell: gõ auto-convert-test-report hoặc actr
# Chạy một lần:  powershell -ExecutionPolicy Bypass -File .\Register-TestReportCommand.ps1

param(
    [switch] $Uninstall
)

$ErrorActionPreference = "Stop"
$Menu = Join-Path $PSScriptRoot "TestReportConvert-Menu.ps1"
$Marker = "AutoConvertTestReport-Khoi"

if (-not (Test-Path -LiteralPath $Menu)) {
    Write-Error "Không thấy: $Menu"
    exit 1
}

# Đường dẫn literal trong function (dấu ' trong path → nhân đôi)
$menuQ = $Menu -replace "'", "''"
$code = @"
# --- $Marker ---
function auto-convert-test-report {
  & '$menuQ'
}
function actr { auto-convert-test-report }
# --- end $Marker ---
"@

if (-not (Test-Path -LiteralPath $PROFILE)) {
    $dir = Split-Path -Parent $PROFILE
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    New-Item -ItemType File -Path $PROFILE -Force | Out-Null
    Write-Host ('Created profile: ' + $PROFILE)
}

$raw = [IO.File]::ReadAllText($PROFILE, [Text.UTF8Encoding]::new($true))

if ($Uninstall) {
    if ($raw -notmatch [regex]::Escape($Marker)) {
        Write-Host ('Not found: ' + $Marker)
        exit 0
    }
    $pat = "(?s)(\r?\n# --- $([regex]::Escape($Marker)) ---\r?\n.*?\r?\n# --- end $([regex]::Escape($Marker)) ---\r?\n?)"
    $raw2 = $raw -replace $pat, "`n"
    [IO.File]::WriteAllText($PROFILE, $raw2.TrimStart(), [Text.UTF8Encoding]::new($true))
    Write-Host ('Uninstall OK. Chay: . {0} trong cua so moi' -f $PROFILE) -ForegroundColor Green
    exit 0
}

if ($raw -match [regex]::Escape($Marker)) {
    Write-Host 'Already installed. Use: actr' -ForegroundColor Yellow
    exit 0
}

$nl = if ($raw.EndsWith("`n")) { "" } else { "`n" }
$append = $nl + $code.Trim() + "`n"
[IO.File]::AppendAllText($PROFILE, $append, [Text.UTF8Encoding]::new($true))
Write-Host ('Added to: ' + $PROFILE) -ForegroundColor Green
Write-Host 'New PS window: actr' -ForegroundColor Cyan
Write-Host ('Or: . {0}' -f $PROFILE) -ForegroundColor DarkGray
