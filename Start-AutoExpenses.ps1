# AutoExpenses launcher - double-click "Start AutoExpenses.bat" or run this script.
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$Host.UI.RawUI.WindowTitle = "AutoExpenses"
$Url = "http://127.0.0.1:8000"
$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"
$OutIndex = Join-Path $Root "frontend-next\out\index.html"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host ">> $Message" -ForegroundColor Cyan
}

function Wait-ForKey {
    Write-Host ""
    Write-Host "Press Enter to close..."
    try {
        [void][System.Console]::ReadLine()
    } catch {
        Start-Sleep -Seconds 3
    }
}

function Fail {
    param([string]$Message)
    Write-Host ""
    Write-Host "ERROR: $Message" -ForegroundColor Red
    Wait-ForKey
    exit 1
}

Write-Host ""
Write-Host "  AutoExpenses" -ForegroundColor Green
Write-Host "  Starting dashboard..." -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path $VenvPython)) {
    Fail "Python virtualenv not found at .venv. Run setup from the README first (python -m venv .venv, then pip install)."
}

if (-not (Test-Path (Join-Path $Root ".env"))) {
    Write-Host "WARNING: No .env file in the project root. Some features may fail until you add one." -ForegroundColor Yellow
}

# Build the UI once if the static export is missing
if (-not (Test-Path $OutIndex)) {
    Write-Step "Dashboard not built yet - running npm run build (first time only)..."
    $Npm = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $Npm) {
        Fail "Node.js/npm not found, and frontend-next\out is missing. Install Node.js 20+ or run npm run build inside frontend-next."
    }
    Push-Location (Join-Path $Root "frontend-next")
    try {
        if (-not (Test-Path "node_modules")) {
            Write-Host "Installing frontend dependencies..."
            npm install
            if ($LASTEXITCODE -ne 0) { Fail "npm install failed." }
        }
        npm run build
        if ($LASTEXITCODE -ne 0) { Fail "npm run build failed." }
    }
    finally {
        Pop-Location
    }
    if (-not (Test-Path $OutIndex)) {
        Fail "Build finished but frontend-next\out\index.html is still missing."
    }
}

Write-Step "Opening browser at $Url"
Start-Process $Url

Write-Step "Starting server (close this window or press Ctrl+C to stop)"
Write-Host "  Dashboard: $Url" -ForegroundColor Green
Write-Host ""

& $VenvPython -m uvicorn api:app --host 127.0.0.1 --port 8000
$ExitCode = $LASTEXITCODE

Write-Host ""
if ($null -ne $ExitCode -and $ExitCode -ne 0) {
    Write-Host "Server exited with code $ExitCode." -ForegroundColor Yellow
}
Wait-ForKey
exit $ExitCode
