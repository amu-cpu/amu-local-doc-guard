param(
    [Parameter(Mandatory = $true)]
    [string]$Url,

    [string]$HealthUrl = "",

    [ValidateSet("Browser", "App")]
    [string]$Mode = "Browser"
)

if ($HealthUrl) {
    $deadline = (Get-Date).AddSeconds(45)
    $ready = $false
    while ((Get-Date) -lt $deadline) {
        try {
            $response = Invoke-RestMethod -Uri $HealthUrl -TimeoutSec 2
            if ($response.status -eq "ok") {
                $ready = $true
                break
            }
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }

    if (-not $ready) {
        exit 1
    }
} else {
    Start-Sleep -Seconds 2
}

if ($Mode -eq "Browser") {
    Start-Process $Url
    exit
}

$edge = (Get-Command msedge.exe -ErrorAction SilentlyContinue).Source
if (-not $edge) {
    $edge = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}

if ($edge) {
    Start-Process $edge -ArgumentList "--app=$Url"
    exit
}

$chrome = (Get-Command chrome.exe -ErrorAction SilentlyContinue).Source
if (-not $chrome) {
    $chrome = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}

if ($chrome) {
    Start-Process $chrome -ArgumentList "--app=$Url"
} else {
    Start-Process $Url
}
