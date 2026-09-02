$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $root ".venv\Scripts\python.exe"

if (-not (Test-Path $python)) {
    py -3 -m venv (Join-Path $root ".venv")
}

& $python -m pip install --upgrade pip
& $python -m pip install -r (Join-Path $root "requirements.txt") pyinstaller
& $python -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name "LTO Vault" `
    --icon (Join-Path $root "assets\lto-vault.ico") `
    --add-data "$(Join-Path $root 'src\index.html');." `
    --collect-all webview `
    --hidden-import clr `
    --hidden-import pythonnet `
    --hidden-import openpyxl `
    (Join-Path $root "src\app.py")

Write-Host "Build complete: dist\LTO Vault.exe" -ForegroundColor Green
