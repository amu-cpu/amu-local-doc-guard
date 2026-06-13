$ErrorActionPreference = "Stop"

$PythonExe = $null
$PythonArgs = @()

if (Test-Path "F:\Tools\Python312\python.exe") {
    $PythonExe = "F:\Tools\Python312\python.exe"
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 --version *> $null
    if ($LASTEXITCODE -eq 0) {
        $PythonExe = "py"
        $PythonArgs = @("-3")
    }
}

if (-not $PythonExe -and (Get-Command python -ErrorAction SilentlyContinue)) {
    & python --version *> $null
    if ($LASTEXITCODE -eq 0) {
        $PythonExe = "python"
    }
}

if (-not $PythonExe) {
    throw "Python 3.10+ was not found. Please install Python, then run this script again."
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    & $PythonExe @PythonArgs -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
