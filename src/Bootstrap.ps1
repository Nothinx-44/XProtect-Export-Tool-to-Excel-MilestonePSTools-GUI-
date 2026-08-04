<#
.SYNOPSIS
    Bootstrap interne de Milestone Toolkit. Appele par Launch.bat.
    Ne pas executer directement — utiliser Launch.bat a la racine du projet.
#>

#Requires -Version 5.1

# Version centrale — source unique dans src/Version.ps1
. (Join-Path $PSScriptRoot 'Version.ps1')

# Applique TLS 1.2 des le debut du processus — requis par PowerShell Gallery.
# PowerShell 5.1 utilise TLS 1.0 par defaut, ce qui bloque Install-Module / Save-Module.
[Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Force Bypass au niveau du processus (complementaire au flag de la ligne de commande)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue

# $PSScriptRoot = .../src/  =>  AppRoot = parent = racine du projet
$AppRoot = if ($PSScriptRoot) {
    Split-Path -Parent $PSScriptRoot
} else {
    Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}


# Masquage console via helper partage (Hide-Console / Show-Console)
. (Join-Path $PSScriptRoot 'Core/ConsoleWindow.ps1')
Hide-Console

# Affiche une erreur fatale de maniere VISIBLE. La console etant masquee,
# un simple Read-Host resterait invisible et l'app semblerait se figer. On privilegie
# une MessageBox WPF ; en dernier recours on reaffiche la console.
function Show-FatalError {
    param([string]$Message)
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        [System.Windows.MessageBox]::Show($Message, 'Milestone Toolkit', 'OK', 'Error') | Out-Null
    }
    catch {
        Show-Console
        Write-Host $Message -ForegroundColor Red
        Read-Host 'Appuyez sur Entree pour quitter'
    }
}

if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Show-FatalError "Milestone Toolkit requires Windows."
    exit 1
}

try {
    . (Join-Path $AppRoot 'src/Core/RequiredModules.ps1')
    . (Join-Path $AppRoot 'src/Core/Show-LanguagePicker.ps1')
    . (Join-Path $AppRoot 'src/Core/Show-StartupCheck.ps1')

    # Lire la langue sauvegardee dans config.json (evite de montrer le selecteur a chaque demarrage)
    $_configPath = Join-Path $AppRoot 'config.json'
    $script:Lang = $null
    if (Test-Path $_configPath) {
        try {
            $_cfg = Get-Content $_configPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($_cfg.language -in @('fr', 'en')) { $script:Lang = $_cfg.language }
        } catch {}
    }

    if (-not $script:Lang) {
        $script:Lang = Show-LanguagePicker
        # Sauvegarder la preference pour les prochains demarrages
        try {
            if (Test-Path $_configPath) {
                $_cfg = Get-Content $_configPath -Raw -Encoding UTF8 | ConvertFrom-Json
            } else {
                $_cfg = [PSCustomObject]@{ outputDirectory='./Output'; snapshotQuality=95; csvDelimiter=';'; csvEncoding='UTF8' }
            }
            Add-Member -InputObject $_cfg -NotePropertyName 'language' -NotePropertyValue $script:Lang -Force
            ConvertTo-Json $_cfg -Depth 10 | Set-Content $_configPath -Encoding UTF8
        } catch {}
    }

    . (Join-Path $AppRoot "src/Lang/$script:Lang.ps1")

    $shouldContinue = Show-StartupCheck -AppRoot $AppRoot
}
catch {
    Show-FatalError "Startup error: $_"
    exit 1
}

if (-not $shouldContinue) { exit 0 }

try {
    & (Join-Path $AppRoot 'src/App.ps1') -RootPath $AppRoot -Lang $script:Lang
}
catch {
    Show-FatalError "Fatal error: $_`n`n$($_.ScriptStackTrace)"
    exit 1
}
