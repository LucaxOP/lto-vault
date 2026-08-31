$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $root ".venv\Scripts\python.exe"

if (-not (Test-Path $python)) {
    py -3 -m venv (Join-Path $root ".venv")
    & $python -m pip install -r (Join-Path $root "requirements.txt")
}

& $python (Join-Path $root "src\app.py")
