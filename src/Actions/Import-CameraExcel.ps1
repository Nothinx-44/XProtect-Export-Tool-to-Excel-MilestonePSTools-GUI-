<#
.SYNOPSIS
    Import de cameras en masse depuis un fichier Excel : generation du MODELE a remplir,
    puis IMPORT (scan auto-detection du pilote -> ajout -> configuration des flux).
.DESCRIPTION
    Le modele contient une feuille "Cameras" (en-tetes fixes + une ligne d'exemple) avec
    des listes deroulantes verrouillees, et une feuille "Reference" (correspondance des
    serveurs, roles de flux, notes). Les en-tetes de colonnes et les valeurs des listes
    sont FIXES (non traduits) pour que l'import soit fiable quelle que soit la langue.

    Import : chaque ligne est scannee via Start-VmsHardwareScan (Milestone auto-detecte le
    pilote), ajoutee via Add-VmsHardware -HardwareScan, puis ses flux sont configures selon
    les colonnes Stream1/2/3 (roles) avec une qualite calculee : Enregistrement = resolution
    max + 20 fps, Live = proche 1080p, Autre = proche 360/400p. La camera doit etre joignable.
    Le reglage fin de la qualite est "best-effort" (les valeurs autorisees dependent du modele).
#>

# En-tetes de colonnes : IDENTIFIANTS FIXES (l'import lira ces noms tels quels).
$script:CamXlsxHeaders = @('Adresse','Nom','Utilisateur','MotDePasse','HTTPS','Serveur','Stream1','Stream2','Stream3')
# Valeurs verrouillees des listes deroulantes (fixes, lues telles quelles a l'import).
$script:CamXlsxHttps   = @('Oui','Non')
$script:CamXlsxRoles   = @('Enregistrement','Live','Les deux','Autre','Aucun')

# Resout le dossier du module ImportExcel (session, catalogue, ou Dependencies/).
function Resolve-ImportExcelBase {
    $mod = Get-Module -Name ImportExcel
    if (-not $mod) { $mod = Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1 }
    if (-not $mod -and $script:DependenciesPath) {
        $psd1 = Join-Path $script:DependenciesPath 'importexcel\ImportExcel.psd1'
        if (Test-Path $psd1) { try { Import-Module $psd1 -Force -ErrorAction Stop; $mod = Get-Module -Name ImportExcel } catch {} }
    }
    if ($mod) { return $mod.ModuleBase }
    return $null
}

# Lance powershell.exe -File <ps1> <json> SANS bloquer le thread UI : pompe la file du
# dispatcher pendant l'attente (fenetre reactive) et met les chemins entre guillemets
# (robuste si %TEMP% contient une espace). Renvoie le process (ExitCode dispo).
function Start-PumpingProcess {
    param([Parameter(Mandatory)] [string]$Ps1, [Parameter(Mandatory)] [string]$JsonArg, [int]$TimeoutSec = 300)
    $argLine = "-ExecutionPolicy Bypass -NonInteractive -File `"$Ps1`" `"$JsonArg`""
    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -PassThru -WindowStyle Hidden
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $proc.HasExited) {
        # Borne de securite : si le sous-processus se bloque (xlsx verrouille, deadlock
        # EPPlus), on le tue plutot que de figer l'UI indefiniment (ExitCode != 0 -> erreur).
        if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { try { $proc.Kill() } catch {} ; break }
        try { [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{}) } catch {}
        Start-Sleep -Milliseconds 120
    }
    return $proc
}

# Ecrit le fichier modele via un sous-processus PowerShell propre (meme approche que
# l'export Hardware : EPPlus depuis le contexte WPF+MilestonePSTools est instable).
function New-CameraExcelTemplate {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter()] [string[]]$ServerNames = @(),
        [Parameter(Mandatory)] [scriptblock]$Log
    )

    $imex = Resolve-ImportExcelBase
    if (-not $imex) { & $Log $script:T.ICT_NoExcel ; return $false }

    # Valeurs de la colonne Serveur (1..n) + correspondance pour la feuille Reference
    $serverValues = @()
    $serverRefLines = @()
    if ($ServerNames.Count -le 1) {
        $serverValues = @('1')
        $only = if ($ServerNames.Count -eq 1) { $ServerNames[0] } else { '?' }
        $serverRefLines = @("  1 = $only", "  ($($script:T.ICT_RefServerSingle))")
    }
    else {
        for ($i = 0; $i -lt $ServerNames.Count; $i++) {
            $serverValues += "$($i + 1)"
            $serverRefLines += "  $($i + 1) = $($ServerNames[$i])"
        }
    }

    # Lignes de la feuille Reference (localisees). "## " = titre de section (stylise).
    $refLines = @("## $($script:T.ICT_RefServersHdr)") + $serverRefLines + @(
        '',
        "## $($script:T.ICT_RefRolesHdr)",
        "  Enregistrement  =  $($script:T.ICT_RoleRecord)",
        "  Live            =  $($script:T.ICT_RoleLive)",
        "  Les deux        =  $($script:T.ICT_RoleBoth)",
        "  Autre           =  $($script:T.ICT_RoleOther)",
        "  Aucun           =  $($script:T.ICT_RoleNone)",
        '',
        "## $($script:T.ICT_RefNotesHdr)",
        "  - $($script:T.ICT_NoteReachable)",
        "  - $($script:T.ICT_NotePassword)"
    )

    $example = @('192.168.1.50', ($script:T.ICT_ExampleName), 'admin', '', 'Non', $serverValues[0], 'Enregistrement', 'Live', 'Autre')

    $payload = [ordered]@{
        Path         = $Path
        ImExPath     = $imex
        SheetCameras = $script:T.ICT_SheetCameras
        SheetRef     = $script:T.ICT_SheetRef
        Headers      = $script:CamXlsxHeaders
        Example      = $example
        HttpsValues  = $script:CamXlsxHttps
        ServerValues = $serverValues
        RoleValues   = $script:CamXlsxRoles
        RefLines     = $refLines
        BannerTitle  = $script:T.ICT_BannerTitle
        BannerSub    = $script:T.ICT_BannerSub
        RefBanner    = $script:T.ICT_RefBanner
        ErrTitle     = $script:T.ICT_ErrValTitle
        ErrMsg       = $script:T.ICT_ErrValMsg
    } | ConvertTo-Json -Depth 4

    $spJson = Join-Path $env:TEMP "MCT_$(Get-Random).json"
    $spPs1  = Join-Path $env:TEMP "MCT_$(Get-Random).ps1"
    Set-Content -Path $spJson -Value $payload -Encoding UTF8

    $spScript = @'
param([string]$J)
try {
    $d = Get-Content $J -Raw -Encoding UTF8 | ConvertFrom-Json
    Import-Module (Join-Path $d.ImExPath 'ImportExcel.psd1') -Force
    Add-Type -AssemblyName System.Drawing
    $Fill = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $Thin = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $HC   = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $VC   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    function C([int]$r,[int]$g,[int]$b){ [System.Drawing.Color]::FromArgb($r,$g,$b) }
    $cBase=C 30 30 46; $cMantle=C 24 24 37; $cSurf=C 49 50 68; $cSurf2=C 69 71 90
    $cText=C 205 214 244; $cSub=C 147 153 178; $cGreen=C 166 227 161; $cBlue=C 137 220 235
    $cRowA=C 255 255 255; $cRowB=C 245 245 250; $cBorder=C 224 224 232; $cExRow=C 245 243 255; $cExTxt=C 127 132 156

    $headers = @($d.Headers)
    $example = @($d.Example)
    $n = $headers.Count
    $row = [ordered]@{}
    for ($i = 0; $i -lt $n; $i++) { $row[$headers[$i]] = $example[$i] }

    # Donnees a partir de la ligne 3 (bandeau titre au-dessus)
    $pkg = ([pscustomobject]$row) | Export-Excel -Path $d.Path -WorksheetName $d.SheetCameras -StartRow 3 -PassThru
    $ws  = $pkg.Workbook.Worksheets[$d.SheetCameras]
    $lastCol = [char]([int][char]'A' + $n - 1)

    # --- Bandeau titre (ligne 1) + sous-titre (ligne 2) ---
    $ws.Cells["A1:${lastCol}1"].Merge = $true
    $ws.Cells["A1"].Value = $d.BannerTitle
    $ws.Cells["A1"].Style.Font.Size = 15; $ws.Cells["A1"].Style.Font.Bold = $true
    $ws.Cells["A1"].Style.Font.Color.SetColor($cGreen)
    $ws.Cells["A1:${lastCol}1"].Style.Fill.PatternType = $Fill
    $ws.Cells["A1:${lastCol}1"].Style.Fill.BackgroundColor.SetColor($cBase)
    $ws.Cells["A1"].Style.VerticalAlignment = $VC; $ws.Cells["A1"].Style.Indent = 1
    $ws.Row(1).Height = 30

    $ws.Cells["A2:${lastCol}2"].Merge = $true
    $ws.Cells["A2"].Value = $d.BannerSub
    $ws.Cells["A2"].Style.Font.Size = 10; $ws.Cells["A2"].Style.Font.Italic = $true
    $ws.Cells["A2"].Style.Font.Color.SetColor($cSub)
    $ws.Cells["A2:${lastCol}2"].Style.Fill.PatternType = $Fill
    $ws.Cells["A2:${lastCol}2"].Style.Fill.BackgroundColor.SetColor($cMantle)
    $ws.Cells["A2"].Style.VerticalAlignment = $VC; $ws.Cells["A2"].Style.Indent = 1
    $ws.Row(2).Height = 18

    # --- En-tete (ligne 3) ---
    $hdr = $ws.Cells["A3:${lastCol}3"]
    $hdr.Style.Fill.PatternType = $Fill; $hdr.Style.Fill.BackgroundColor.SetColor($cSurf)
    $hdr.Style.Font.Bold = $true; $hdr.Style.Font.Color.SetColor($cText)
    $hdr.Style.HorizontalAlignment = $HC; $hdr.Style.VerticalAlignment = $VC
    $ws.Row(3).Height = 22

    # --- Ligne d'exemple (ligne 4) ---
    $ex = $ws.Cells["A4:${lastCol}4"]
    $ex.Style.Font.Italic = $true; $ex.Style.Font.Color.SetColor($cExTxt)
    $ex.Style.Fill.PatternType = $Fill; $ex.Style.Fill.BackgroundColor.SetColor($cExRow)

    # --- Zebrage + bordures des lignes de saisie (4..40) ---
    for ($r = 5; $r -le 40; $r++) {
        $rng = $ws.Cells["A${r}:${lastCol}${r}"]
        $rng.Style.Fill.PatternType = $Fill
        $rng.Style.Fill.BackgroundColor.SetColor($(if ($r % 2 -eq 0) { $cRowA } else { $cRowB }))
    }
    for ($r = 4; $r -le 40; $r++) {
        $b = $ws.Cells["A${r}:${lastCol}${r}"].Style.Border.Bottom
        $b.Style = $Thin; $b.Color.SetColor($cBorder)
    }
    $tbl = $ws.Cells["A3:${lastCol}40"]
    foreach ($bd in @($tbl.Style.Border.Top,$tbl.Style.Border.Bottom,$tbl.Style.Border.Left,$tbl.Style.Border.Right)) {
        $bd.Style = $Thin; $bd.Color.SetColor($cSurf2)
    }

    # Largeurs de colonnes
    $widths = 16,22,15,15,9,10,15,15,15
    for ($i = 0; $i -lt $n; $i++) { $ws.Column($i + 1).Width = $widths[$i] }

    # Filtre + volets figes
    $ws.Cells["A3:${lastCol}3"].AutoFilter = $true
    $ws.View.FreezePanes(4, 1)

    # --- Validations (lignes 4..1003) ---
    function Add-Val($ws, $range, $vals, $title, $msg) {
        $v = $ws.DataValidations.AddListValidation($range)
        foreach ($x in $vals) { [void]$v.Formula.Values.Add([string]$x) }
        $v.ShowErrorMessage = $true
        $v.ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::stop
        $v.ErrorTitle = $title; $v.Error = $msg
    }
    Add-Val $ws 'E4:E1003' @($d.HttpsValues)  $d.ErrTitle $d.ErrMsg
    Add-Val $ws 'F4:F1003' @($d.ServerValues) $d.ErrTitle $d.ErrMsg
    Add-Val $ws 'G4:I1003' @($d.RoleValues)   $d.ErrTitle $d.ErrMsg

    # --- Feuille Reference ---
    $refWs = $pkg.Workbook.Worksheets.Add($d.SheetRef)
    $refWs.Column(1).Width = 82
    $refWs.Cells["A1"].Value = $d.RefBanner
    $refWs.Cells["A1"].Style.Font.Size = 14; $refWs.Cells["A1"].Style.Font.Bold = $true
    $refWs.Cells["A1"].Style.Font.Color.SetColor($cGreen)
    $refWs.Cells["A1"].Style.Fill.PatternType = $Fill
    $refWs.Cells["A1"].Style.Fill.BackgroundColor.SetColor($cBase)
    $refWs.Cells["A1"].Style.VerticalAlignment = $VC; $refWs.Cells["A1"].Style.Indent = 1
    $refWs.Row(1).Height = 28

    $lines = @($d.RefLines)
    $rr = 3
    foreach ($ln in $lines) {
        $txt = [string]$ln
        if ($txt -like '## *') {
            $refWs.Cells[$rr,1].Value = $txt.Substring(3)
            $refWs.Cells[$rr,1].Style.Font.Bold = $true
            $refWs.Cells[$rr,1].Style.Font.Color.SetColor($cBlue)
            $refWs.Cells[$rr,1].Style.Fill.PatternType = $Fill
            $refWs.Cells[$rr,1].Style.Fill.BackgroundColor.SetColor($cSurf)
            $refWs.Row($rr).Height = 20
        } else {
            $refWs.Cells[$rr,1].Value = $txt
            $refWs.Cells[$rr,1].Style.Font.Color.SetColor((C 60 60 70))
        }
        $rr++
    }

    $pkg.Save(); $pkg.Dispose()
} finally { Remove-Item $J -Force -ErrorAction SilentlyContinue }
'@
    Set-Content -Path $spPs1 -Value $spScript -Encoding UTF8

    try {
        if (Test-Path $Path) { Remove-Item $Path -Force -ErrorAction Stop }
    } catch { & $Log ($script:T.ICT_Err -f $script:T.ICT_FileLocked) ; return $false }

    $ok = $false
    try {
        $proc = Start-PumpingProcess -Ps1 $spPs1 -JsonArg $spJson
        if ($proc.ExitCode -eq 0 -and (Test-Path $Path)) { $ok = $true }
        else { & $Log ($script:T.ICT_Err -f "code $($proc.ExitCode)") }
    }
    catch { & $Log ($script:T.ICT_Err -f $_.Exception.Message) }
    finally {
        Remove-Item $spPs1  -Force -ErrorAction SilentlyContinue
        Remove-Item $spJson -Force -ErrorAction SilentlyContinue   # filet si le sous-processus n'a pas demarre
    }

    return $ok
}


# ============================================================
# IMPORT depuis Excel
# ============================================================

# Normalise le libelle de role d'un flux vers une action canonique.
# '' (cellule vide) -> 'skip' (ne touche pas au flux). Valeur inconnue -> 'skip' aussi.
function Convert-StreamRole {
    param([string]$Raw)
    $r = "$Raw".Trim().ToLowerInvariant()
    if (-not $r) { return 'skip' }
    switch -Regex ($r) {
        '^(enregistrement|enreg|record|rec)$' { return 'record' }
        '^(live|direct)$'                     { return 'live' }
        '^(les deux|les 2|both|deux)$'        { return 'both' }
        '^(autre|other|secondaire|low)$'      { return 'other' }
        '^(aucun|off|none|desactive|non)$'    { return 'off' }
        default                               { return 'skip' }
    }
}

# Cles d'un objet de reglages (hashtable/dictionnaire ou pscustomobject).
function Get-SettingKeys {
    param($Obj)
    if ($null -eq $Obj) { return @() }
    if ($Obj -is [System.Collections.IDictionary]) { return @($Obj.Keys) }
    return @($Obj.PSObject.Properties.Name)
}

# Extrait, best-effort, la liste des valeurs autorisees d'un reglage depuis la sortie
# -ValueTypeInfo (forme variable selon le SDK : on teste plusieurs structures).
function Get-AllowedValues {
    param($Info, [string]$Key)
    if ($null -eq $Info) { return @() }
    try {
        $entry = $null
        if ($Info -is [System.Collections.IDictionary]) { $entry = $Info[$Key] }
        elseif ($Info.PSObject.Properties[$Key])        { $entry = $Info.$Key }
        if ($null -eq $entry) {
            $entry = $Info | Where-Object { "$($_.Key)$($_.Name)" -match [regex]::Escape($Key) } | Select-Object -First 1
        }
        if ($null -eq $entry) { return @() }
        foreach ($p in 'Values','AllowedValues','ValueList','Options','ValueTypeInfos') {
            if ($entry.PSObject.Properties[$p] -and $entry.$p) {
                return @($entry.$p | ForEach-Object { if ($_.PSObject.Properties['Value']) { "$($_.Value)" } else { "$_" } })
            }
        }
        if ($entry -is [System.Collections.IEnumerable] -and $entry -isnot [string]) {
            return @($entry | ForEach-Object { if ($_.PSObject.Properties['Value']) { "$($_.Value)" } else { "$_" } })
        }
    } catch {}
    return @()
}

# Choisit une resolution ("LxH") parmi les valeurs autorisees selon la cible.
function Select-Resolution {
    param([string[]]$Allowed, [string]$Target)
    $parsed = @()
    foreach ($a in $Allowed) {
        $m = [regex]::Match("$a", '(\d{2,5})\s*[xX*]\s*(\d{2,5})')
        if ($m.Success) { $parsed += [pscustomobject]@{ Raw = $a; W = [int]$m.Groups[1].Value; H = [int]$m.Groups[2].Value } }
    }
    if ($parsed.Count -eq 0) { return $null }
    switch ($Target) {
        'max'  { return ($parsed | Sort-Object { $_.W * $_.H } -Descending      | Select-Object -First 1).Raw }
        '1080' { return ($parsed | Sort-Object { [math]::Abs($_.H - 1080) }      | Select-Object -First 1).Raw }
        '360'  { return ($parsed | Sort-Object { [math]::Abs($_.H - 380) }       | Select-Object -First 1).Raw }
    }
    return $null
}

# Choisit la valeur numerique la plus proche d'une cible (ex. FPS = 20).
function Select-NearestNumber {
    param([string[]]$Allowed, [double]$Target)
    $nums = @()
    foreach ($a in $Allowed) {
        $m = [regex]::Match("$a", '(\d+(\.\d+)?)')
        if ($m.Success) { $nums += [pscustomobject]@{ Raw = $a; N = [double]$m.Groups[1].Value } }
    }
    if ($nums.Count -eq 0) { return $null }
    return ($nums | Sort-Object { [math]::Abs($_.N - $Target) } | Select-Object -First 1).Raw
}

# Applique la qualite (resolution / fps) a un flux — best-effort, ne leve jamais.
function Set-StreamQualityBestEffort {
    param($Camera, [string]$StreamName, [string]$ResTarget, $Fps, [scriptblock]$Log)
    try {
        $cur  = Get-VmsDeviceStreamSetting -Device $Camera -StreamName $StreamName -ErrorAction Stop
        if ($cur -is [array]) { $cur = $cur[0] }   # certains drivers renvoient un tableau
        $info = $null
        try { $info = Get-VmsDeviceStreamSetting -Device $Camera -StreamName $StreamName -ValueTypeInfo -ErrorAction Stop } catch {}
        $resKey = $null; $fpsKey = $null
        foreach ($k in (Get-SettingKeys $cur)) {
            if (-not $resKey -and $k -match '(?i)resolution')                 { $resKey = $k }
            if (-not $fpsKey -and $k -match '(?i)(fps|framerate|frame ?rate)') { $fpsKey = $k }
        }
        $settings = @{}
        if ($resKey) {
            $pick = Select-Resolution (Get-AllowedValues $info $resKey) $ResTarget
            if ($pick) { $settings[$resKey] = $pick }
        }
        if ($Fps -and $fpsKey) {
            $pickF = Select-NearestNumber (Get-AllowedValues $info $fpsKey) ([double]$Fps)
            if ($pickF) { $settings[$fpsKey] = $pickF }
        }
        if ($settings.Count -gt 0) {
            Set-VmsDeviceStreamSetting -Device $Camera -StreamName $StreamName -Settings $settings -ErrorAction Stop
            & $Log ($script:T.IC_Quality -f $StreamName, (($settings.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '))
        }
    }
    catch { & $Log ($script:T.IC_QualityWarn -f $StreamName, $_.Exception.Message) }
}

# Configure les flux 1/2/3 d'un materiel selon les roles issus de l'Excel.
function Set-CameraStreamsFromRow {
    param($Hardware, [string[]]$Roles, [scriptblock]$Log)
    $cams = @()
    try { $cams = @($Hardware | Get-VmsCamera) } catch {}
    if ($cams.Count -eq 0) { & $Log $script:T.IC_NoCam ; return }
    foreach ($cam in $cams) {
        $streams = @()
        try { $streams = @(Get-VmsCameraStream -Camera $cam) } catch {}
        for ($i = 0; $i -lt 3; $i++) {
            $role = Convert-StreamRole $Roles[$i]
            if ($role -eq 'skip') { continue }
            if ($i -ge $streams.Count) { & $Log ($script:T.IC_NoStream -f ($i + 1)) ; continue }
            $s = $streams[$i]
            try {
                switch ($role) {
                    'off'    { Set-VmsCameraStream -Stream $s -Disabled -ErrorAction Stop | Out-Null }
                    'record' { Set-VmsCameraStream -Stream $s -Recorded -ErrorAction Stop | Out-Null
                               Set-StreamQualityBestEffort $cam "$($s.Name)" 'max' 20 $Log }
                    'live'   { Set-VmsCameraStream -Stream $s -LiveDefault -ErrorAction Stop | Out-Null
                               Set-StreamQualityBestEffort $cam "$($s.Name)" '1080' $null $Log }
                    'both'   { Set-VmsCameraStream -Stream $s -Recorded -LiveDefault -ErrorAction Stop | Out-Null
                               Set-StreamQualityBestEffort $cam "$($s.Name)" 'max' 20 $Log }
                    'other'  { Set-StreamQualityBestEffort $cam "$($s.Name)" '360' $null $Log }
                }
            }
            catch { & $Log ($script:T.IC_StreamErr -f ($i + 1), $_.Exception.Message) }
        }
    }
}

# Lit les lignes de la feuille "Cameras" via un sous-processus (EPPlus propre).
function Read-CameraExcelRows {
    param([Parameter(Mandatory)] [string]$Path, [Parameter(Mandatory)] [scriptblock]$Log)
    $imex = Resolve-ImportExcelBase
    if (-not $imex) { & $Log $script:T.ICT_NoExcel ; return $null }

    $outJson = Join-Path $env:TEMP "MCR_out_$(Get-Random).json"
    $payload = [ordered]@{ Xlsx = $Path; Sheet = $script:T.ICT_SheetCameras; ImExPath = $imex; Out = $outJson } | ConvertTo-Json
    $spJson  = Join-Path $env:TEMP "MCR_$(Get-Random).json"
    $spPs1   = Join-Path $env:TEMP "MCR_$(Get-Random).ps1"
    Set-Content -Path $spJson -Value $payload -Encoding UTF8

    $spScript = @'
param([string]$J)
try {
    $d = Get-Content $J -Raw -Encoding UTF8 | ConvertFrom-Json
    Import-Module (Join-Path $d.ImExPath 'ImportExcel.psd1') -Force
    $info  = Get-ExcelSheetInfo $d.Xlsx
    $sheet = if (($info | ForEach-Object Name) -contains $d.Sheet) { $d.Sheet } else { $info[0].Name }
    # Detecte la ligne d'en-tete (celle contenant "Adresse") : le modele a un bandeau
    # au-dessus, donc l'en-tete n'est pas forcement en ligne 1.
    $hr = 1
    $probe = @(Import-Excel -Path $d.Xlsx -WorksheetName $sheet -NoHeader -EndRow 8)
    for ($k = 0; $k -lt $probe.Count; $k++) {
        if (@($probe[$k].PSObject.Properties.Value) -contains 'Adresse') { $hr = $k + 1; break }
    }
    $rows  = @(Import-Excel -Path $d.Xlsx -WorksheetName $sheet -StartRow $hr)
    $out   = $rows | Select-Object Adresse,Nom,Utilisateur,MotDePasse,HTTPS,Serveur,Stream1,Stream2,Stream3
    @($out) | ConvertTo-Json -Depth 4 | Set-Content -Path $d.Out -Encoding UTF8
} finally { Remove-Item $J -Force -ErrorAction SilentlyContinue }
'@
    Set-Content -Path $spPs1 -Value $spScript -Encoding UTF8

    $rows = $null
    try {
        $proc = Start-PumpingProcess -Ps1 $spPs1 -JsonArg $spJson
        if ($proc.ExitCode -eq 0 -and (Test-Path $outJson)) {
            $raw = Get-Content $outJson -Raw -Encoding UTF8
            if ($raw) { $rows = @($raw | ConvertFrom-Json) }
        }
        else { & $Log ($script:T.IC_ErrRead -f "code $($proc.ExitCode)") }
    }
    catch { & $Log ($script:T.IC_ErrRead -f $_.Exception.Message) }
    finally {
        Remove-Item $spPs1   -Force -ErrorAction SilentlyContinue
        Remove-Item $spJson  -Force -ErrorAction SilentlyContinue   # filet si le sous-processus n'a pas demarre
        Remove-Item $outJson -Force -ErrorAction SilentlyContinue
    }
    return $rows
}

# Construit un SecureString depuis une chaine, y compris VIDE (contrairement a
# ConvertTo-SecureString -AsPlainText qui leve une erreur sur une valeur vide).
function ConvertTo-SecureStringSafe {
    param([string]$Plain)
    $ss = New-Object System.Security.SecureString
    foreach ($ch in "$Plain".ToCharArray()) { $ss.AppendChar($ch) }
    $ss.MakeReadOnly()
    return $ss
}

# Scan auto-detection d'une adresse (Milestone trouve le pilote seul) puis ajout.
# Partage par l'ajout manuel ET l'import Excel. Renvoie @{ Hardware; Status; Message } :
#   Status = 'added' | 'exists' (deja present localement) | 'notfound' (rien de valide).
# Les erreurs dures (scan / ajout) sont propagees a l'appelant (bloc try/catch).
function Add-CameraByScan {
    param(
        [Parameter(Mandatory)] $RecordingServer,
        [Parameter(Mandatory)] [uri]$Uri,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential]$Credential,
        [bool]$UseHttps = $false
    )
    $scanParams = @{ RecordingServer = $RecordingServer; Address = $Uri; Credential = $Credential; PassThru = $true; ErrorAction = 'Stop' }
    if ($UseHttps) { $scanParams.UseHttps = $true }
    $scan  = @(Start-VmsHardwareScan @scanParams)
    $valid = @($scan | Where-Object { $_.HardwareScanValidated })
    if ($valid.Count -eq 0) {
        $reason = if ($scan.Count -gt 0 -and $scan[0].ErrorText) { "$($scan[0].ErrorText)" } else { $script:T.IC_NoDetect }
        return @{ Hardware = $null; Status = 'notfound'; Message = $reason }
    }
    $target = @($valid | Where-Object { -not $_.MacAddressExistsLocal }) | Select-Object -First 1
    if (-not $target) { return @{ Hardware = $null; Status = 'exists'; Message = '' } }
    $hw = $target | Add-VmsHardware -Force -ErrorAction Stop
    return @{ Hardware = $hw; Status = 'added'; Message = '' }
}

# Construit l'URI d'une camera. Un schema explicite dans l'adresse prime sur le defaut
# HTTPS. Renvoie @{ Uri=[uri] ou $null si invalide ; Https=[bool] }. Partage manuel/import.
function Resolve-CameraUri {
    param([Parameter(Mandatory)] [string]$Address, [bool]$DefaultHttps = $false)
    $addr  = "$Address".Trim()
    $https = $DefaultHttps
    if     ($addr -match '^(?i)https://') { $https = $true ; $uriText = $addr }
    elseif ($addr -match '^(?i)http://')  { $https = $false; $uriText = $addr }
    elseif ($addr -match '^[a-zA-Z][a-zA-Z0-9+.\-]*://') { $uriText = $addr }
    elseif ($https)                        { $uriText = "https://$addr" }
    else                                   { $uriText = "http://$addr" }
    $uri = $null
    [void][System.Uri]::TryCreate($uriText, [System.UriKind]::Absolute, [ref]$uri)
    return @{ Uri = $uri; Https = $https }
}

# Renomme un materiel si un nom non vide est fourni (echec silencieux, best-effort).
# LogKey = cle $script:T du message de succes (AC_LogRenamed / IC_Renamed).
function Rename-HardwareSafe {
    param([Parameter(Mandatory)] $Hardware, [string]$Name, [scriptblock]$Log = {}, [string]$LogKey)
    $n = "$Name".Trim()
    if (-not $n) { return }
    try {
        Set-VmsHardware -Hardware $Hardware -Name $n -ErrorAction Stop | Out-Null
        if ($LogKey) { & $Log ($script:T.$LogKey -f $n) }
    } catch {}
}

# Orchestration de l'import : lecture -> scan -> ajout -> configuration des flux.
function Import-CameraExcelRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {}
    )

    if (-not (Test-Path $Path)) { & $Log ($script:T.IC_FileNotFound -f $Path) ; return }

    & $Log $script:T.IC_Reading
    $rows = Read-CameraExcelRows -Path $Path -Log $Log
    if (-not $rows -or @($rows).Count -eq 0) { & $Log $script:T.IC_NoRows ; return }

    $servers = @()
    try { $servers = @(Get-VmsRecordingServer) } catch {}
    if ($servers.Count -eq 0) { & $Log $script:T.AC_ValNoServer ; return }

    # Garde les lignes avec une adresse, en excluant la ligne d'exemple du modele
    # (nom contenant "(exemple)" / "(example)") pour ne pas generer de faux echec.
    $items = @($rows | Where-Object {
        "$($_.Adresse)".Trim() -and ("$($_.Nom)" -notmatch '(?i)\((exemple|example)\)')
    })
    $total = $items.Count
    if ($total -eq 0) { & $Log $script:T.IC_NoRows ; return }

    & $Log ($script:T.IC_Start -f $total)
    $done = 0; $ok = 0; $err = 0; $skip = 0
    foreach ($r in $items) {
        if (& $Cancel) { break }
        $done++
        & $ReportProgress $done $total
        $addr = "$($r.Adresse)".Trim()
        & $Log ($script:T.IC_Row -f $done, $total, $addr)

        try {
            # Serveur : colonne "1/2/3" -> index ; vide -> premier serveur
            $srvIdx = 0
            $sv = "$($r.Serveur)".Trim()
            if ($sv -match '^\d+$') { $srvIdx = [int]$sv - 1 }
            if ($srvIdx -lt 0 -or $srvIdx -ge $servers.Count) { $srvIdx = 0 }
            $rs = $servers[$srvIdx]

            $user = "$($r.Utilisateur)".Trim()
            if (-not $user) { $err++; & $Log ($script:T.IC_RowErr -f $addr, $script:T.AC_ValNoUser) ; continue }
            $cred = New-Object System.Management.Automation.PSCredential($user, (ConvertTo-SecureStringSafe "$($r.MotDePasse)"))

            $httpsCol = ("$($r.HTTPS)".Trim() -match '^(?i)(oui|yes|true|1|o|y)$')
            $u = Resolve-CameraUri -Address $addr -DefaultHttps $httpsCol
            if (-not $u.Uri) { $err++; & $Log ($script:T.IC_RowErr -f $addr, $script:T.IC_BadAddress) ; continue }

            # Scan auto-detection du pilote + ajout (logique partagee)
            $res = Add-CameraByScan -RecordingServer $rs -Uri $u.Uri -Credential $cred -UseHttps $u.Https
            if ($res.Status -eq 'notfound') { $err++;  & $Log ($script:T.IC_RowErr -f $addr, $res.Message) ; continue }
            if ($res.Status -eq 'exists')   { $skip++; & $Log ($script:T.IC_Exists -f $addr) ; continue }
            $newHw = $res.Hardware
            $ok++
            & $Log ($script:T.IC_RowOk -f $addr)

            # Post-ajout (renommage + flux) : le materiel est deja cree, une erreur ici
            # ne doit PAS le recompter en echec.
            try {
                Rename-HardwareSafe -Hardware $newHw -Name "$($r.Nom)" -Log $Log -LogKey 'IC_Renamed'
                Set-CameraStreamsFromRow -Hardware $newHw -Roles @("$($r.Stream1)", "$($r.Stream2)", "$($r.Stream3)") -Log $Log
            }
            catch { & $Log ($script:T.IC_StreamErr -f '-', $_.Exception.Message) }
        }
        catch { $err++; & $Log ($script:T.IC_RowErr -f $addr, $_.Exception.Message) }
    }
    if ($skip -gt 0) { & $Log ($script:T.IC_Skipped -f $skip) }
    & $Log ($script:T.IC_Done -f $ok, $err)
}
