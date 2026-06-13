$ErrorActionPreference = "Stop"
$utf8 = [Text.UTF8Encoding]::new()
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

function T([string]$Base64) {
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Base64))
}

$projectDir = Split-Path -Parent $PSScriptRoot
$targetPath = Join-Path $projectDir (T "5ZCv5Yqo5bel5YW3X+W6lOeUqOeql+WPoy5iYXQ=")
$iconPath = Join-Path $projectDir "app.ico"
$desktopDir = [Environment]::GetFolderPath("DesktopDirectory")
$shortcutPath = Join-Path $desktopDir (T "5ZWG5Lia6K6h5YiS5Lmm6ISx5pWP6K+m5oOF5Zu+5bel5YW3Lmxuaw==")

if (-not (Test-Path $targetPath)) {
    throw "Missing target: $targetPath"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
$shortcut.WorkingDirectory = $projectDir
$shortcut.Description = T "5ZWG5Lia6K6h5YiS5Lmm6ISx5pWP6K+m5oOF5Zu+55Sf5oiQ5bel5YW3"
if (Test-Path $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
}
$shortcut.Save()

Write-Host (T "5qGM6Z2i5b+r5o235pa55byP5bey5Yib5bu677yM5LmL5ZCO5Y+v55u05o6l5Y+M5Ye75qGM6Z2i5Zu+5qCH5ZCv5Yqo5bel5YW344CC")
Write-Host ("{0}{1}" -f (T "5b+r5o235pa55byP5L2N572u77ya"), $shortcutPath)
