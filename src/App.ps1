<#
.SYNOPSIS
    Point d'entree principal de l'application Milestone Toolkit.
#>

param(
    [Parameter()]
    [string]$RootPath,
    [Parameter()]
    [string]$Lang = 'fr'
)

$script:AppRoot    = $RootPath
$script:SrcPath    = Join-Path $AppRoot 'src'

# Version : source unique dans src/Version.ps1 (charge AVANT la langue qui l'utilise)
. (Join-Path $script:SrcPath 'Version.ps1')

# Chargement de la langue
. (Join-Path $script:SrcPath "Lang/$Lang.ps1")

$configPath = Join-Path $AppRoot 'config.json'
if (Test-Path $configPath) {
    $configRaw = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
else {
    $configRaw = [PSCustomObject]@{}
}

# Valeurs par defaut appliquees CLE PAR CLE : un config.json partiel (edite a la main
# ou issu d'une ancienne version) ne doit pas produire snapshotQuality = 0 (cast [int]
# de $null) ni un csvEncoding $null qui fait echouer Export-Csv.
$configDefaults = @{
    outputDirectory = './Output'
    snapshotQuality = 95
    csvDelimiter    = ';'
    csvEncoding     = 'UTF8'
}
foreach ($k in $configDefaults.Keys) {
    $existing = $configRaw.PSObject.Properties[$k]
    if (-not $existing -or [string]::IsNullOrWhiteSpace("$($existing.Value)")) {
        Add-Member -InputObject $configRaw -NotePropertyName $k -NotePropertyValue $configDefaults[$k] -Force
    }
}

$outputDir = $configRaw.outputDirectory
if (-not [System.IO.Path]::IsPathRooted($outputDir)) {
    $outputDir = Join-Path $AppRoot $outputDir
}

$script:DependenciesPath = Join-Path $AppRoot 'Dependencies'

# snapshotQuality : valide 1..100, sinon retour au defaut (une valeur non numerique
# dans le JSON ferait sinon planter le cast [int] au demarrage).
$snapQuality = 0
if (-not [int]::TryParse("$($configRaw.snapshotQuality)", [ref]$snapQuality) -or
    $snapQuality -lt 1 -or $snapQuality -gt 100) {
    $snapQuality = $configDefaults.snapshotQuality
}

$script:Config = @{
    outputDirectory = $outputDir
    snapshotQuality = $snapQuality
    csvDelimiter    = $configRaw.csvDelimiter
    csvEncoding     = $configRaw.csvEncoding
    logDirectory    = Join-Path $AppRoot 'Logs'
}

# ============================================================
# CHARGEMENT DES SCRIPTS
# ============================================================

. (Join-Path $SrcPath 'Core/ConsoleWindow.ps1')
. (Join-Path $SrcPath 'Core/RequiredModules.ps1')
. (Join-Path $SrcPath 'Core/Initialize-Modules.ps1')
. (Join-Path $SrcPath 'Core/Write-ActivityLog.ps1')
. (Join-Path $SrcPath 'Core/Invoke-PtzPreset.ps1')

. (Join-Path $SrcPath 'Actions/Get-SnapshotSelected.ps1')
. (Join-Path $SrcPath 'Actions/Get-SnapshotAll.ps1')
. (Join-Path $SrcPath 'Actions/Export-HardwareReport.ps1')
. (Join-Path $SrcPath 'Actions/Set-CameraGroupByModel.ps1')
. (Join-Path $SrcPath 'Actions/New-BulkAlarm.ps1')
. (Join-Path $SrcPath 'Actions/Add-CameraDevice.ps1')
. (Join-Path $SrcPath 'Actions/Import-CameraExcel.ps1')
. (Join-Path $SrcPath 'Actions/Get-PtzPresetSnapshot.ps1')
. (Join-Path $SrcPath 'Actions/Get-RecordingStats.ps1')
. (Join-Path $SrcPath 'Actions/Get-LicenseInfo.ps1')
. (Join-Path $SrcPath 'Actions/Get-CameraStatus.ps1')
. (Join-Path $SrcPath 'Actions/Get-PlaybackReport.ps1')

# Purge des anciens logs (> 30 jours) — une fois au demarrage
Remove-OldLogs -LogDirectory $script:Config.logDirectory -RetentionDays 30

# ============================================================
# INITIALISATION MODULES ET CONNEXION
# ============================================================

if (Test-Path $DependenciesPath) {
    $installMode = 'Offline'
    Write-Host 'Offline mode detected.' -ForegroundColor Yellow
}
else {
    $installMode = 'Online'
}

$initLog = { param($Message) Write-Host $Message }
Initialize-RequiredModules -InstallMode $installMode -DependenciesPath $DependenciesPath -Log $initLog

Write-Host 'Connecting to Milestone server...' -ForegroundColor Cyan
try {
    $connectParams = @{ ShowDialog = $true; AcceptEula = $true; Force = $true }
    # Auto-login du dialogue Milestone DESACTIVE par defaut : si un ancien serveur
    # (inaccessible) a ete memorise avec "Auto login", le dialogue se connecte tout
    # seul et l'application plante au demarrage sans laisser changer de serveur.
    # Reactivable en ajoutant "autoLogin": true dans config.json.
    if ($configRaw.autoLogin -ne $true) {
        $connectParams.DisableAutoLogin = $true
    }
    Connect-ManagementServer @connectParams
    Write-Host 'Connected.' -ForegroundColor Green
}
catch {
    # La console est masquee : sans MessageBox, l'echec de connexion serait invisible.
    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
    [System.Windows.MessageBox]::Show(
        ($script:T.App_ConnectFailed -f $_.Exception.Message),
        $script:T.App_ConnectFailedTitle, 'OK', 'Error') | Out-Null
    exit 1
}

# ============================================================
# MASQUER LA CONSOLE
# ============================================================

Hide-Console

# ============================================================
# CHARGEMENT WPF
# ============================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$xamlPath    = Join-Path $SrcPath 'UI/MainWindow.xaml'
$xamlContent = Get-Content $xamlPath -Raw -Encoding UTF8
$xamlReader  = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlContent))
$script:Window = [System.Windows.Markup.XamlReader]::Load($xamlReader)

# ============================================================
# REFERENCES UI
# ============================================================

$script:LogOutput        = $Window.FindName('LogOutput')
$script:ActionStatus     = $Window.FindName('ActionStatus')
$script:ProgressBar      = $Window.FindName('ProgressBar')
$script:StatusIndicator  = $Window.FindName('StatusIndicator')
$script:StatusText       = $Window.FindName('StatusText')
$script:OutputDirText    = $Window.FindName('OutputDirText')

$script:BtnSnapshotSelected = $Window.FindName('BtnSnapshotSelected')
$script:BtnSnapshotAll      = $Window.FindName('BtnSnapshotAll')
$script:BtnPtzSnapshot      = $Window.FindName('BtnPtzSnapshot')
$script:BtnExportHardware   = $Window.FindName('BtnExportHardware')
$script:BtnGroupByModel     = $Window.FindName('BtnGroupByModel')
$script:BtnAlarms           = $Window.FindName('BtnAlarms')
$script:BtnAddCamera        = $Window.FindName('BtnAddCamera')
$script:BtnRecordingStats   = $Window.FindName('BtnRecordingStats')
$script:BtnLicenseInfo      = $Window.FindName('BtnLicenseInfo')
$script:BtnCameraStatus     = $Window.FindName('BtnCameraStatus')
$script:BtnPlaybackReport   = $Window.FindName('BtnPlaybackReport')
$script:BtnClearLog         = $Window.FindName('BtnClearLog')
$script:BtnCancel           = $Window.FindName('BtnCancel')
$script:BtnOutputDir        = $Window.FindName('BtnOutputDir')
$script:SnapshotMode        = $Window.FindName('SnapshotMode')
$script:DateTimePanel       = $Window.FindName('DateTimePanel')
$script:SnapshotDate        = $Window.FindName('SnapshotDate')
$script:SnapshotHour        = $Window.FindName('SnapshotHour')
$script:SnapshotMinute      = $Window.FindName('SnapshotMinute')

# ============================================================
# APPLICATION DES TEXTES TRADUITS
# ============================================================

$Window.Title                               = $script:T.MW_AppTitle
$Window.FindName('LblOutputDirHeader').Text = $script:T.MW_LblOutputDir
$Window.FindName('LblVersion').Text         = $script:T.MW_Version
$Window.FindName('LblModeCapture').Text     = $script:T.MW_LblModeCapture
$Window.FindName('LblDate').Text            = $script:T.MW_LblDate
$Window.FindName('LblHeure').Text           = $script:T.MW_LblHeure
$Window.FindName('LblSnapshots').Text       = $script:T.MW_LblSnapshots
$Window.FindName('LblGestion').Text         = $script:T.MW_LblGestion
$Window.FindName('LblMonitoring').Text      = $script:T.MW_LblMonitoring
$Window.FindName('LblDiagnostic').Text      = $script:T.MW_LblDiagnostic
$Window.FindName('LblJournal').Text         = $script:T.MW_LblJournal

$script:BtnOutputDir.Content = $script:T.MW_BtnOutputDir
$script:BtnClearLog.Content  = $script:T.MW_BtnClearLog
$script:BtnCancel.Content    = $script:T.MW_BtnCancel

$script:SnapshotMode.Items[0].Content = $script:T.MW_CbiLive
$script:SnapshotMode.Items[1].Content = $script:T.MW_CbiHistorique

$script:BtnSnapshotSelected.Content.Children[1].Text = $script:T.MW_BtnSnapshotSel
$script:BtnSnapshotAll.Content.Children[1].Text      = $script:T.MW_BtnSnapshotAll
$script:BtnPtzSnapshot.Content.Children[1].Text      = $script:T.MW_BtnPtz
$script:BtnExportHardware.Content.Children[1].Text   = $script:T.MW_BtnExportHardware
$script:BtnGroupByModel.Content.Children[1].Text     = $script:T.MW_BtnGroupByModel
$script:BtnAlarms.Content.Children[1].Text           = $script:T.MW_BtnAlarms
$script:BtnAddCamera.Content.Children[1].Text        = $script:T.MW_BtnAddCamera
$script:BtnCameraStatus.Content.Children[1].Text     = $script:T.MW_BtnCameraStatus
$script:BtnPlaybackReport.Content.Children[1].Text   = $script:T.MW_BtnPlaybackReport
$script:BtnRecordingStats.Content.Children[1].Text   = $script:T.MW_BtnRecordingStats
$script:BtnLicenseInfo.Content.Children[1].Text      = $script:T.MW_BtnLicenseInfo

# Init document RichTextBox
$script:LogOutput.Document.Blocks.Clear()
$script:LogOutput.Document.PagePadding = [System.Windows.Thickness]::new(16, 12, 16, 12)

$script:StatusIndicator.Fill = [System.Windows.Media.Brushes]::LightGreen
$script:StatusText.Text      = $script:T.MW_StatusConnected
$script:OutputDirText.Text   = $script:Config.outputDirectory
$script:SnapshotDate.SelectedDate = [datetime]::Today.AddDays(-1)

# ============================================================
# ETAT PARTAGE
# ============================================================

$script:CancelRequested = $false
$script:IsCancelled = { $script:CancelRequested }
$script:LastUIPump  = 0

# Force le traitement de la file du dispatcher (rendu + entrees souris).
# Permet au bouton Annuler de s'afficher et de reagir pendant une action.
function Update-UI {
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Input, [Action]{}
    )
}

# Retourne @{ Ok=$true; Time=[datetime] ou $null } — type uniforme, plus de $false sentinel
function Get-SnapshotDateTime {
    if ($script:SnapshotMode.SelectedIndex -eq 0) {
        return @{ Ok = $true; Time = $null }
    }

    $date = $script:SnapshotDate.SelectedDate
    if (-not $date) {
        [System.Windows.MessageBox]::Show(
            $script:T.App_DateMissing, $script:T.App_DateTitle, 'OK', 'Warning') | Out-Null
        return @{ Ok = $false; Time = $null }
    }

    $h = 0; $m = 0
    if (-not [int]::TryParse($script:SnapshotHour.Text, [ref]$h) -or $h -lt 0 -or $h -gt 23) {
        [System.Windows.MessageBox]::Show(
            $script:T.App_HourInvalid, $script:T.App_HourTitle, 'OK', 'Warning') | Out-Null
        return @{ Ok = $false; Time = $null }
    }
    if (-not [int]::TryParse($script:SnapshotMinute.Text, [ref]$m) -or $m -lt 0 -or $m -gt 59) {
        [System.Windows.MessageBox]::Show(
            $script:T.App_MinInvalid, $script:T.App_MinTitle, 'OK', 'Warning') | Out-Null
        return @{ Ok = $false; Time = $null }
    }

    return @{ Ok = $true; Time = $date.Date.AddHours($h).AddMinutes($m) }
}

$script:ReportProgress = {
    param([int]$Current, [int]$Total)
    if ($Total -gt 0) {
        $script:ProgressBar.IsIndeterminate = $false
        $script:ProgressBar.Maximum         = [double]$Total
        $script:ProgressBar.Value           = [double]$Current
    }
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background, [Action]{}
    )
}

# ============================================================
# FONCTIONS UI
# ============================================================

function Write-UILog {
    param([string]$Message, [string]$Level = 'INFO')

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $fullMsg   = "[$timestamp] $Message"

    $brush = switch ($Level) {
        'ERROR'   { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(243,139,168)) }
        'WARN'    { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(249,168, 37)) }
        'SUCCESS' { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(166,227,161)) }
        'ACTION'  { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(137,180,250)) }
        default   { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(205,214,244)) }
    }

    $run  = [System.Windows.Documents.Run]::new($fullMsg)
    $run.Foreground = $brush
    $para = [System.Windows.Documents.Paragraph]::new($run)
    $para.Margin     = [System.Windows.Thickness]::new(0)
    $para.LineHeight = 20
    $script:LogOutput.Document.Blocks.Add($para)
    if ($script:LogOutput.Document.Blocks.Count -gt 2000) {
        $script:LogOutput.Document.Blocks.Remove($script:LogOutput.Document.Blocks.FirstBlock)
    }
    $script:LogOutput.ScrollToEnd()

    # Pump du dispatcher limite (~50 ms) : rafraichit l'UI sans payer un rendu
    # synchrone a chaque ligne, ce qui ralentissait les actions verbeuses.
    $now = [Environment]::TickCount
    if (($now - $script:LastUIPump) -ge 50) {
        $script:LastUIPump = $now
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Render, [Action]{}
        )
    }

    $fileLevel = if ($Level -in 'ACTION','INFO') { 'INFO' } else { $Level }
    Write-ActivityLog -Message $Message -Level $fileLevel -LogDirectory $script:Config.logDirectory
}

$script:ActionButtons = @(
    $BtnSnapshotSelected, $BtnSnapshotAll, $BtnPtzSnapshot,
    $BtnExportHardware, $BtnGroupByModel, $BtnAlarms, $BtnAddCamera,
    $BtnCameraStatus, $BtnPlaybackReport,
    $BtnRecordingStats, $BtnLicenseInfo
)

function Set-UIBusy {
    param([string]$ActionName)
    foreach ($btn in $script:ActionButtons) { $btn.IsEnabled = $false }
    $script:BtnCancel.Visibility        = [System.Windows.Visibility]::Visible
    $script:ActionStatus.Text           = $ActionName
    $script:ProgressBar.IsIndeterminate = $true
    $script:ProgressBar.Visibility      = [System.Windows.Visibility]::Visible
    $Window.Cursor                      = [System.Windows.Input.Cursors]::Wait
}

function Set-UIReady {
    foreach ($btn in $script:ActionButtons) { $btn.IsEnabled = $true }
    $script:BtnCancel.Visibility        = [System.Windows.Visibility]::Collapsed
    $script:ActionStatus.Text           = $script:T.MW_StatusReady
    $script:ProgressBar.IsIndeterminate = $false
    $script:ProgressBar.Value           = 0
    $script:ProgressBar.Visibility      = [System.Windows.Visibility]::Collapsed
    $Window.Cursor                      = [System.Windows.Input.Cursors]::Arrow
}

function Invoke-Action {
    param([string]$Name, [scriptblock]$Action)

    $script:CancelRequested = $false
    Set-UIBusy -ActionName $Name
    Write-UILog "--- $Name ---" 'ACTION'
    # Garantit que l'etat occupe (bouton Annuler visible) est rendu et hit-testable
    # AVANT de demarrer le travail. La reactivite pendant l'action depend ensuite des
    # appels a ReportProgress/log qui pompent la file du dispatcher.
    Update-UI

    try {
        & $Action
        if ($script:CancelRequested) {
            Write-UILog $script:T.App_ActionCancelled 'WARN'
        }
        else {
            Write-UILog $script:T.App_ActionDone 'SUCCESS'
        }
    }
    catch {
        Write-UILog "ERREUR: $_" 'ERROR'
        Write-ActivityLog -Message $_.Exception.Message -Level 'ERROR' -LogDirectory $script:Config.logDirectory
    }
    finally {
        Set-UIReady
    }
}

# ============================================================
# EVENEMENTS
# ============================================================

$logCallback = {
    param([string]$Message)
    $trimmed = $Message.TrimStart()
    $level = if ($trimmed -match '^ERREUR|^ERROR')           { 'ERROR' }
             elseif ($trimmed -match '^AVERTISSEMENT|^WARN') { 'WARN'  }
             else                                            { 'INFO'  }
    Write-UILog -Message $Message -Level $level
}

$script:SnapshotMode.Add_SelectionChanged({
    $vis = if ($script:SnapshotMode.SelectedIndex -eq 1) { 'Visible' } else { 'Collapsed' }
    $script:DateTimePanel.Visibility = $vis
})

$BtnSnapshotSelected.Add_Click({
    $snap = Get-SnapshotDateTime
    if (-not $snap.Ok) { return }
    $script:SnapshotTime = $snap.Time
    Invoke-Action -Name $script:T.Act_SnapshotSel -Action {
        Get-SnapshotSelected -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -SnapshotTime $script:SnapshotTime
    }
})

$BtnSnapshotAll.Add_Click({
    $snap = Get-SnapshotDateTime
    if (-not $snap.Ok) { return }
    $script:SnapshotTime = $snap.Time
    Invoke-Action -Name $script:T.Act_SnapshotAll -Action {
        Get-SnapshotAll -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress `
            -SnapshotTime $script:SnapshotTime
    }
})

$BtnPtzSnapshot.Add_Click({
    $snap = Get-SnapshotDateTime
    if (-not $snap.Ok) { return }
    $script:SnapshotTime = $snap.Time
    Invoke-Action -Name $script:T.Act_SnapshotPtz -Action {
        Get-PtzPresetSnapshot -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress `
            -SnapshotTime $script:SnapshotTime
    }
})

$BtnExportHardware.Add_Click({
    Invoke-Action -Name $script:T.Act_ExportHW -Action {
        Export-HardwareReport -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnGroupByModel.Add_Click({
    Invoke-Action -Name $script:T.Act_GroupModel -Action {
        Set-CameraGroupByModel -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnAlarms.Add_Click({
    Invoke-Action -Name $script:T.Act_Alarms -Action {
        New-BulkAlarm -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnAddCamera.Add_Click({
    Invoke-Action -Name $script:T.Act_AddCamera -Action {
        Add-CameraDevice -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnCameraStatus.Add_Click({
    Invoke-Action -Name $script:T.Act_CamStatus -Action {
        Get-CameraStatus -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnPlaybackReport.Add_Click({
    Invoke-Action -Name $script:T.Act_Playback -Action {
        Get-PlaybackReport -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnRecordingStats.Add_Click({
    Invoke-Action -Name $script:T.Act_RecStats -Action {
        Get-RecordingStats -Config $script:Config -Log $logCallback `
            -Cancel $script:IsCancelled -ReportProgress $script:ReportProgress
    }
})

$BtnLicenseInfo.Add_Click({
    Invoke-Action -Name $script:T.Act_License -Action {
        Get-VmsLicenseSummary -Log $logCallback
    }
})

$BtnClearLog.Add_Click({
    $script:LogOutput.Document.Blocks.Clear()
})

$BtnCancel.Add_Click({
    $script:CancelRequested = $true
    $script:BtnCancel.IsEnabled = $false
    $script:ActionStatus.Text   = $script:T.MW_StatusCancelling
})

$BtnOutputDir.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description       = $script:T.App_ChooseDir
    $dialog.SelectedPath      = $script:Config.outputDirectory
    $dialog.ShowNewFolderButton = $true

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Config.outputDirectory = $dialog.SelectedPath
        $script:OutputDirText.Text     = $dialog.SelectedPath
        Write-UILog ($script:T.App_OutputChanged -f $dialog.SelectedPath) 'INFO'
    }
})

$Window.Add_Closing({
    Write-ActivityLog -Message $script:T.App_Closing -Level 'INFO' -LogDirectory $script:Config.logDirectory
    try { Disconnect-ManagementServer } catch {}
    Show-Console
})

# ============================================================
# AFFICHAGE
# ============================================================

Write-UILog $script:T.App_Started 'SUCCESS'
Write-UILog ($script:T.App_OutputDir -f $script:Config.outputDirectory)

$Window.Add_Loaded({ $Window.Activate() })
[void]$Window.ShowDialog()
