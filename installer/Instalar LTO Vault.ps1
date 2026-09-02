$ErrorActionPreference = "Stop"

$source = Join-Path $PSScriptRoot "LTO Vault.exe"
$installDirectory = Join-Path $env:LOCALAPPDATA "Programs\LTO Vault"
$installedApp = Join-Path $installDirectory "LTO Vault.exe"

if (-not (Test-Path -LiteralPath $source)) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("O arquivo LTO Vault.exe deve estar na mesma pasta do instalador.", "LTO Vault") | Out-Null
    exit 1
}

try {
    New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $installedApp -Force

    $shell = New-Object -ComObject WScript.Shell
    $desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "LTO Vault.lnk"
    $startMenuDirectory = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\LTO Vault"
    New-Item -ItemType Directory -Path $startMenuDirectory -Force | Out-Null

    foreach ($shortcutPath in @($desktopShortcut, (Join-Path $startMenuDirectory "LTO Vault.lnk"))) {
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $installedApp
        $shortcut.WorkingDirectory = $installDirectory
        $shortcut.IconLocation = "$installedApp,0"
        $shortcut.Description = "Controle local de backups em fitas LTO"
        $shortcut.Save()
    }

    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("LTO Vault 1.0 foi instalado. Um atalho foi criado na área de trabalho.", "Instalação concluída") | Out-Null
}
catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("Não foi possível instalar o LTO Vault.`n`n$($_.Exception.Message)", "Erro na instalação") | Out-Null
    exit 1
}
