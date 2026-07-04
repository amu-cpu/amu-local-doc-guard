param(
    [ValidateSet("Browser", "App")]
    [string]$Mode = "Browser"
)

$ErrorActionPreference = "Stop"
$utf8 = [Text.UTF8Encoding]::new()
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONIOENCODING = "utf-8"

function Stop-WithMessage([string]$Message) {
    Write-Host $Message
    Read-Host "Press Enter to exit..." | Out-Null
    exit 1
}

$projectDir = Split-Path -Parent $PSScriptRoot
$outputDir = Join-Path $projectDir "output"
$pythonF = "F:\Tools\Python312\python.exe"
$venvPython = Join-Path $projectDir ".venv\Scripts\python.exe"
$activateScript = Join-Path $projectDir ".venv\Scripts\Activate.ps1"
$requirements = Join-Path $projectDir "requirements.txt"
$selectPortScript = Join-Path $projectDir "scripts\select_port.ps1"
$openPageScript = Join-Path $projectDir "scripts\open_local_page.ps1"

Set-Location -LiteralPath $projectDir

Write-Host "========================================"
Write-Host "Local desensitizer detail-image tool"
if ($Mode -eq "App") {
    Write-Host "App window mode"
}
Write-Host "========================================"
Write-Host ("Project directory: " + $projectDir)
Write-Host ("Output directory: " + $outputDir)

if (-not (Test-Path -LiteralPath $pythonF)) {
    Write-Host ("Python not found: " + $pythonF)
    Stop-WithMessage "Please install Python 3.10 or later."
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Project virtual environment not found. Creating .venv ..."
    & $pythonF -m venv (Join-Path $projectDir ".venv")
    if ($LASTEXITCODE -ne 0) {
        Stop-WithMessage "Startup failed. Please check the error above."
    }
}

if (Test-Path -LiteralPath $activateScript) {
    . $activateScript
} else {
    Write-Host ("Virtual environment activate script not found. Using Python directly: " + $venvPython)
}

Write-Host "Installing / checking dependencies..."
& $venvPython -m pip install --no-cache-dir -r $requirements
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "Dependency installation failed."
}

$port = (& powershell -NoProfile -ExecutionPolicy Bypass -File $selectPortScript | Select-Object -First 1)
if (-not $port) {
    Stop-WithMessage "Ports 5000 and 5001 are both occupied."
}

$env:APP_PORT = [string]$port
$env:APP_DEBUG = "0"
$localUrl = "http://127.0.0.1:$port/"
$healthUrl = "http://127.0.0.1:$port/health"
$currentPython = (& $venvPython -c "import sys; print(sys.executable)")
$pythonVersion = (& $venvPython --version)

Write-Host ("Current Python path: " + $currentPython)
Write-Host "Current Python version:"
Write-Host $pythonVersion
Write-Host ("Local URL: " + $localUrl)
Write-Host ("Project directory: " + $projectDir)
Write-Host ("Output directory: " + $outputDir)
Write-Host ""
Write-Host "Keep this window open while using the tool."
Write-Host "Waiting for local service to start..."
Write-Host "Starting local service. Please do not close this window."

Start-Process powershell -WindowStyle Hidden -ArgumentList @(
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    $openPageScript,
    "-Url",
    $localUrl,
    "-HealthUrl",
    $healthUrl,
    "-Mode",
    $Mode
)

& $venvPython (Join-Path $projectDir "app.py")
if ($LASTEXITCODE -ne 0) {
    Stop-WithMessage "Startup failed. Please check the error above."
}
