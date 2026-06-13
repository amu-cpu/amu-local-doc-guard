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

function T([string]$Base64) {
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Base64))
}

function Say([string]$Base64) {
    Write-Host (T $Base64)
}

function SayValue([string]$LabelBase64, [string]$Value) {
    Write-Host ("{0}{1}" -f (T $LabelBase64), $Value)
}

function StopWithMessage([string]$Base64) {
    Say $Base64
    Read-Host (T "5oyJ5Lu75oSP6ZSu6YCA5Ye6Li4u") | Out-Null
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

Set-Location $projectDir

Write-Host "========================================"
Say "5ZWG5Lia6K6h5YiS5LmmIC8g5Y+v56CU5oql5ZGK6ISx5pWP6K+m5oOF5Zu+55Sf5oiQ5bel5YW3"
if ($Mode -eq "App") {
    Say "5bqU55So56qX5Y+j5qih5byP"
}
Write-Host "========================================"
SayValue "5b2T5YmN6aG555uu55uu5b2V77ya" $projectDir
SayValue "6L6T5Ye655uu5b2V5L2N572u77ya" $outputDir

if (-not (Test-Path $pythonF)) {
    SayValue "6ZSZ6K+v5o+Q56S677ya5pyq5om+5YiwIFB5dGhvbu+8mg==" $pythonF
    StopWithMessage "6K+35YWI5a6J6KOFIFB5dGhvbiAzLjEwIOaIluabtOmrmOeJiOacrOOAgg=="
}

if (-not (Test-Path $venvPython)) {
    Say "5pyq5om+5Yiw6aG555uu6Jma5ouf546v5aKD77yM5q2j5Zyo5Yib5bu6IC52ZW52IC4uLg=="
    & $pythonF -m venv (Join-Path $projectDir ".venv")
    if ($LASTEXITCODE -ne 0) {
        StopWithMessage "6ZSZ6K+v5o+Q56S677ya5ZCv5Yqo5aSx6LSl77yM6K+35p+l55yL5LiK6Z2i55qE6ZSZ6K+v5L+h5oGv44CC"
    }
}

. $activateScript

Say "5q2j5Zyo5a6J6KOFL+ajgOafpemhueebruS+nei1li4uLg=="
& $venvPython -m pip install --no-cache-dir -r $requirements
if ($LASTEXITCODE -ne 0) {
    StopWithMessage "6ZSZ6K+v5o+Q56S677ya5L6d6LWW5a6J6KOF5aSx6LSl44CC"
}

$port = (& powershell -NoProfile -ExecutionPolicy Bypass -File $selectPortScript | Select-Object -First 1)
if (-not $port) {
    StopWithMessage "6ZSZ6K+v5o+Q56S677yaNTAwMCDlkowgNTAwMSDnq6/lj6Ppg73ooqvljaDnlKjjgII="
}

$env:APP_PORT = [string]$port
$env:APP_DEBUG = "0"
$localUrl = "http://127.0.0.1:$port/"
$healthUrl = "http://127.0.0.1:$port/health"
$currentPython = (& $venvPython -c "import sys; print(sys.executable)")
$pythonVersion = (& $venvPython --version)

SayValue "5b2T5YmNIFB5dGhvbiDot6/lvoTvvJo=" $currentPython
Say "5b2T5YmNIFB5dGhvbiDniYjmnKzvvJo="
Write-Host $pythonVersion
SayValue "5b2T5YmN6K6/6Zeu5Zyw5Z2A77ya" $localUrl
SayValue "5b2T5YmN6aG555uu55uu5b2V77ya" $projectDir
SayValue "6L6T5Ye655uu5b2V5L2N572u77ya" $outputDir
Write-Host ""
Say "6K+35LiN6KaB5YWz6Zet5q2k56qX5Y+j77yM5YWz6Zet5ZCO572R6aG15bCG5peg5rOV6K6/6Zeu"
Say "5q2j5Zyo562J5b6F5pys5Zyw5pyN5Yqh5ZCv5YqoLi4u"
Say "5q2j5Zyo5ZCv5Yqo5pys5Zyw5pyN5Yqh77yM6K+35LiN6KaB5YWz6Zet5q2k56qX5Y+j44CC"

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
    StopWithMessage "6ZSZ6K+v5o+Q56S677ya5ZCv5Yqo5aSx6LSl77yM6K+35p+l55yL5LiK6Z2i55qE6ZSZ6K+v5L+h5oGv44CC"
}
