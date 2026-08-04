<#
.SYNOPSIS
    Applique une mise a jour de Milestone Toolkit apres que le processus principal a quitte.
    Appele par le bouton de mise a jour dans Show-StartupCheck — ne pas executer directement.
.PARAMETER Target
    Dossier racine de l'installation a mettre a jour.
.PARAMETER Source
    Dossier contenant les nouveaux fichiers extraits depuis l'archive GitHub.
.PARAMETER Launcher
    Chemin du fichier .bat a relancer apres la mise a jour.
#>
param(
    [Parameter(Mandatory)] [string]$Target,
    [Parameter(Mandatory)] [string]$Source,
    [Parameter(Mandatory)] [string]$Launcher
)

$errLog = Join-Path $env:TEMP 'MilestoneToolkitUpdate_error.txt'
function Write-UpdErr([string]$m) {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $m" | Out-File -FilePath $errLog -Encoding UTF8 -Append
}

# Attend que le processus parent libere les fichiers (max 5 s)
Start-Sleep -Seconds 2
for ($i = 0; $i -lt 20; $i++) {
    try   { Get-ChildItem -Path $Target -ErrorAction Stop | Out-Null; break }
    catch { Start-Sleep -Milliseconds 250 }
}

# Dossiers a NE PAS sauvegarder (donnees utilisateur / volumineux). On ne les ecrase
# pas non plus : la copie de la mise a jour n'y touche que si l'archive les contient.
$excludeDirs = @('Dependencies', 'Output', 'Logs', '.git')

# Sauvegarde du code actuel avant ecrasement (permet un rollback si la copie echoue).
$backupDir = Join-Path $env:TEMP ("MilestoneToolkitBackup_{0:yyyyMMdd_HHmmss}" -f (Get-Date))
$backupOk  = $false
try {
    New-Item -Path $backupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    Get-ChildItem -Path $Target -Force | Where-Object { $_.Name -notin $excludeDirs } |
        Copy-Item -Destination $backupDir -Recurse -Force -ErrorAction Stop
    $backupOk = $true
}
catch {
    Write-UpdErr "Sauvegarde impossible (mise a jour poursuivie sans rollback) : $_"
}

try {
    Copy-Item -Path (Join-Path $Source '*') -Destination $Target -Recurse -Force -ErrorAction Stop
    # Les nouveaux fichiers peuvent porter le "Mark of the Web" : on supprime le temoin
    # pour que le .bat les debloque au prochain lancement.
    Remove-Item (Join-Path $Target '.unblocked') -Force -ErrorAction SilentlyContinue
    # Succes : la sauvegarde n'est plus necessaire.
    if ($backupOk) { Remove-Item $backupDir -Recurse -Force -ErrorAction SilentlyContinue }
}
catch {
    Write-UpdErr "Echec de la copie des fichiers de mise a jour : $_"
    if ($backupOk) {
        try {
            Get-ChildItem -Path $backupDir -Force |
                Copy-Item -Destination $Target -Recurse -Force -ErrorAction Stop
            Write-UpdErr 'Restauration de la version precedente reussie.'
        }
        catch {
            Write-UpdErr "ECHEC CRITIQUE de la restauration : $_ (sauvegarde conservee dans $backupDir)"
        }
    }
}

if (Test-Path $Launcher) {
    Start-Process -FilePath $Launcher
}
