<#
.SYNOPSIS
    Ajout de materiel / cameras Milestone par AUTO-DETECTION du pilote, avec clonage
    optionnel des reglages de flux d'un modele existant.
.DESCRIPTION
    Ouvre une fenetre avec :
      - Mode "Une camera" (une adresse) ou "Liste" (plusieurs adresses, un par ligne).
      - Identifiants du peripherique (mot de passe en SecureString, jamais journalise).
      - Option "Cloner les reglages de flux d'un modele existant" : reprend codec /
        resolution / FPS / qualite d'un materiel deja present (le pilote, lui, reste
        auto-detecte par Milestone).
    L'ajout passe par Start-VmsHardwareScan (detection du pilote) puis Add-VmsHardware
    -HardwareScan -- logique partagee Add-CameraByScan (cf. Import-CameraExcel.ps1). Le
    clonage lit les flux du modele via Get-VmsCameraStream / Set-VmsCameraStream.

    IMPORTANT : Milestone exige que la camera soit JOIGNABLE sur le reseau pour l'ajouter
    (le serveur d'enregistrement se connecte a l'appareil). Il n'existe aucun moyen
    supporte d'ajouter du materiel injoignable (echec VMO60281 sinon).
#>

# Modele lisible d'un materiel existant (fallback sur le nom si Model absent).
function Get-HardwareModelLabel {
    param([object]$Hardware)
    $m = "$($Hardware.Model)"
    if ([string]::IsNullOrWhiteSpace($m)) { $m = "$($Hardware.Name)" }
    return $m
}

function Show-CameraAddDialog {
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Windows.Forms   -ErrorAction SilentlyContinue

    # --- Serveurs d'enregistrement ---
    $recServers = @()
    try { $recServers = @(Get-VmsRecordingServer) } catch {}
    if ($recServers.Count -eq 0) {
        [System.Windows.MessageBox]::Show($script:T.AC_ValNoServer, $script:T.AC_Title, 'OK', 'Error') | Out-Null
        return $null
    }

    # --- Materiel existant : source des modeles a cloner (dedupe par modele) ---
    $existingHw = @()
    try { $existingHw = @(Get-VmsHardware) } catch {}
    $templates = [System.Collections.Generic.List[object]]::new()
    $seenModels = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    # Label calcule une seule fois par materiel (tri + dedup)
    $labeled = $existingHw | ForEach-Object { [pscustomobject]@{ Hw = $_; Model = (Get-HardwareModelLabel $_) } }
    foreach ($x in ($labeled | Sort-Object Model)) {
        if ($seenModels.Add($x.Model)) {
            $templates.Add([pscustomobject]@{ Model = $x.Model; Hardware = $x.Hw })
        }
    }

    $script:_AC_Result = $null

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($script:T.AC_Title)"
        Width="560" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#1E1E2E" FontFamily="Segoe UI">
    <Window.Resources>
        <Style x:Key="FlatBtn" TargetType="Button">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#45475A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <StackPanel Margin="20,16,20,18">
        <TextBlock Text="$($script:T.AC_Header)" Foreground="#CDD6F4" FontSize="18" FontWeight="Bold"/>
        <TextBlock Text="$($script:T.AC_Subtitle)" Foreground="#9399B2" FontSize="12"
                   TextWrapping="Wrap" Margin="0,4,0,10"/>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="12,10" Margin="0,0,0,12">
            <StackPanel>
                <TextBlock Text="$($script:T.AC_XlsHeader)" Foreground="#89DCEB" FontSize="12"
                           FontWeight="Bold" Margin="0,0,0,6"/>
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="BtnXlsImport" Content="$($script:T.AC_XlsImport)"
                            Style="{StaticResource FlatBtn}"
                            Background="#89DCEB" Foreground="#1E1E2E" BorderThickness="0"
                            FontWeight="Bold" FontSize="12" Padding="16,7" Cursor="Hand" Margin="0,0,10,0"/>
                    <Button x:Name="BtnXlsTemplate" Content="$($script:T.AC_XlsTemplate)"
                            Style="{StaticResource FlatBtn}"
                            Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                            FontSize="12" Padding="16,7" Cursor="Hand"/>
                </StackPanel>
                <TextBlock Text="$($script:T.AC_XlsHint)" Foreground="#9399B2" FontSize="11"
                           TextWrapping="Wrap" Margin="0,6,0,0"/>
            </StackPanel>
        </Border>

        <TextBlock Text="$($script:T.AC_ManualHeader)" Foreground="#CDD6F4" FontSize="12"
                   FontWeight="Bold" Margin="0,0,0,6"/>

        <TextBlock Text="$($script:T.AC_LblRecServer)" Foreground="#CDD6F4" FontSize="12" Margin="0,0,0,3"/>
        <ComboBox x:Name="CboServer" FontSize="12" Margin="0,0,0,12" Padding="6,4"/>

        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="$($script:T.AC_LblMode)" Foreground="#CDD6F4" FontSize="12"
                       VerticalAlignment="Center" Margin="0,0,14,0"/>
            <RadioButton x:Name="RbSingle" Content="$($script:T.AC_ModeSingle)" GroupName="Mode"
                         IsChecked="True" Foreground="#CDD6F4" FontSize="12" Margin="0,0,16,0"/>
            <RadioButton x:Name="RbList" Content="$($script:T.AC_ModeList)" GroupName="Mode"
                         Foreground="#CDD6F4" FontSize="12"/>
        </StackPanel>

        <StackPanel x:Name="PanelSingle">
            <TextBlock Text="$($script:T.AC_LblAddress)" Foreground="#CDD6F4" FontSize="12" Margin="0,0,0,3"/>
            <TextBox x:Name="TxtAddress" FontSize="12" Padding="6,4" Margin="0,0,0,8"
                     Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A"/>
            <TextBlock Text="$($script:T.AC_LblName)" Foreground="#CDD6F4" FontSize="12" Margin="0,0,0,3"/>
            <TextBox x:Name="TxtName" FontSize="12" Padding="6,4" Margin="0,0,0,12"
                     Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A"/>
        </StackPanel>

        <StackPanel x:Name="PanelList" Visibility="Collapsed">
            <TextBlock Text="$($script:T.AC_LblList)" Foreground="#CDD6F4" FontSize="12" Margin="0,0,0,3"/>
            <TextBox x:Name="TxtList" FontSize="12" Padding="6,4" Margin="0,0,0,12" Height="110"
                     AcceptsReturn="True" TextWrapping="NoWrap"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                     Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A"
                     FontFamily="Consolas"/>
        </StackPanel>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="12,10" Margin="0,0,0,12">
            <StackPanel>
                <TextBlock Text="$($script:T.AC_LblCred)" Foreground="#F9E2AF" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="14"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Text="$($script:T.AC_LblUser)" Foreground="#CDD6F4" FontSize="11" Margin="0,0,0,3"/>
                        <TextBox x:Name="TxtUser" FontSize="12" Padding="6,4"
                                 Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Text="$($script:T.AC_LblPassword)" Foreground="#CDD6F4" FontSize="11" Margin="0,0,0,3"/>
                        <PasswordBox x:Name="PwdBox" FontSize="12" Padding="6,4"
                                     Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A"/>
                    </StackPanel>
                </Grid>
                <CheckBox x:Name="ChkHttps" Content="$($script:T.AC_LblHttps)" Foreground="#CDD6F4"
                          FontSize="11" Margin="0,8,0,0"/>
            </StackPanel>
        </Border>

        <CheckBox x:Name="ChkClone" Content="$($script:T.AC_LblCloneStreams)"
                  Foreground="#CDD6F4" FontSize="12" Margin="0,4,0,0"/>
        <StackPanel x:Name="PanelTpl" Margin="20,6,0,6" Visibility="Collapsed">
            <TextBlock Text="$($script:T.AC_LblTemplate)" Foreground="#CDD6F4" FontSize="11" Margin="0,0,0,3"/>
            <ComboBox x:Name="CboTemplate" FontSize="12" Padding="6,4" Margin="0,0,0,4"/>
            <TextBlock Text="$($script:T.AC_TplHint)" Foreground="#9399B2" FontSize="11" TextWrapping="Wrap"/>
        </StackPanel>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button x:Name="BtnCancel" Content="$($script:T.AC_BtnCancel)" IsCancel="True"
                    Style="{StaticResource FlatBtn}"
                    Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                    FontSize="13" Padding="22,8" Cursor="Hand" Margin="0,0,10,0"/>
            <Button x:Name="BtnAdd" Content="$($script:T.AC_BtnAdd)" IsDefault="True"
                    Style="{StaticResource FlatBtn}"
                    Background="#A6E3A1" Foreground="#1E1E2E" BorderThickness="0"
                    FontWeight="Bold" FontSize="13" Padding="26,8" Cursor="Hand"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $win    = [System.Windows.Markup.XamlReader]::Load($reader)

    $CboServer   = $win.FindName('CboServer')
    $RbSingle    = $win.FindName('RbSingle')
    $RbList      = $win.FindName('RbList')
    $PanelSingle = $win.FindName('PanelSingle')
    $PanelList   = $win.FindName('PanelList')
    $TxtAddress  = $win.FindName('TxtAddress')
    $TxtName     = $win.FindName('TxtName')
    $TxtList     = $win.FindName('TxtList')
    $TxtUser     = $win.FindName('TxtUser')
    $PwdBox      = $win.FindName('PwdBox')
    $ChkHttps    = $win.FindName('ChkHttps')
    $PanelTpl    = $win.FindName('PanelTpl')
    $CboTemplate = $win.FindName('CboTemplate')
    $ChkClone    = $win.FindName('ChkClone')
    $BtnAdd      = $win.FindName('BtnAdd')
    $BtnXlsImport   = $win.FindName('BtnXlsImport')
    $BtnXlsTemplate = $win.FindName('BtnXlsTemplate')

    # Noms des serveurs (pour la generation du modele Excel)
    $serverNames = @($recServers | ForEach-Object { "$($_.Name)" })

    # --- Boutons Excel : renvoient une action au lieu de l'ajout manuel ---
    $BtnXlsImport.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title  = $script:T.AC_XlsOpenTitle
        $ofd.Filter = 'Excel (*.xlsx)|*.xlsx'
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:_AC_Result = @{ ExcelAction = 'Import'; Path = $ofd.FileName }
            $win.DialogResult = $true
        }
    })
    $BtnXlsTemplate.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title    = $script:T.ICT_SaveTitle
        $sfd.Filter   = 'Excel (*.xlsx)|*.xlsx'
        $sfd.FileName = $script:T.ICT_FileName
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:_AC_Result = @{ ExcelAction = 'Template'; Path = $sfd.FileName; ServerNames = $serverNames }
            $win.DialogResult = $true
        }
    })

    # --- Serveurs ---
    foreach ($rs in $recServers) {
        $it = New-Object System.Windows.Controls.ComboBoxItem
        $it.Content = "$($rs.Name)"
        $it.Tag     = $rs
        [void]$CboServer.Items.Add($it)
    }
    $CboServer.SelectedIndex = 0

    # --- Modeles a cloner ---
    if ($templates.Count -eq 0) {
        $it = New-Object System.Windows.Controls.ComboBoxItem
        $it.Content   = $script:T.AC_NoTemplates
        $it.IsEnabled = $false
        [void]$CboTemplate.Items.Add($it)
        $ChkClone.IsEnabled = $false   # aucun modele existant a cloner
    }
    else {
        foreach ($t in $templates) {
            $it = New-Object System.Windows.Controls.ComboBoxItem
            $it.Content = $script:T.AC_TplItem -f $t.Model, "$($t.Hardware.Name)"
            $it.Tag     = $t.Hardware
            [void]$CboTemplate.Items.Add($it)
        }
        $CboTemplate.SelectedIndex = 0
    }

    # --- Bascule des panneaux ---
    $RbSingle.Add_Checked({ $PanelSingle.Visibility = 'Visible'; $PanelList.Visibility = 'Collapsed' })
    $RbList.Add_Checked({ $PanelSingle.Visibility = 'Collapsed'; $PanelList.Visibility = 'Visible' })
    # Le pilote est auto-detecte par Milestone : le modele ne sert plus qu'au clonage.
    $ChkClone.Add_Checked({ $PanelTpl.Visibility = 'Visible' })
    $ChkClone.Add_Unchecked({ $PanelTpl.Visibility = 'Collapsed' })

    # --- Validation / construction du resultat ---
    $BtnAdd.Add_Click({
        $rsItem = $CboServer.SelectedItem
        if (-not $rsItem) {
            [System.Windows.MessageBox]::Show($script:T.AC_ValNoServer, $script:T.AC_Title, 'OK', 'Warning') | Out-Null
            return
        }

        # Adresses (mode unique ou liste)
        $items = [System.Collections.Generic.List[object]]::new()
        if ($RbSingle.IsChecked) {
            $addr = "$($TxtAddress.Text)".Trim()
            if ($addr) { $items.Add(@{ Address = $addr; Name = "$($TxtName.Text)".Trim() }) }
        }
        else {
            foreach ($line in ("$($TxtList.Text)" -split "`r?`n")) {
                $l = $line.Trim()
                if (-not $l) { continue }
                $parts = $l -split ';', 2
                $a = $parts[0].Trim()
                if (-not $a) { continue }   # ignore une ligne sans adresse (ex. ";nom")
                $items.Add(@{ Address = $a; Name = if ($parts.Count -gt 1) { $parts[1].Trim() } else { '' } })
            }
        }
        if ($items.Count -eq 0) {
            [System.Windows.MessageBox]::Show($script:T.AC_ValNoAddress, $script:T.AC_Title, 'OK', 'Warning') | Out-Null
            return
        }

        $user = "$($TxtUser.Text)".Trim()
        if ([string]::IsNullOrWhiteSpace($user)) {
            [System.Windows.MessageBox]::Show($script:T.AC_ValNoUser, $script:T.AC_Title, 'OK', 'Warning') | Out-Null
            return
        }

        # Clonage optionnel : le pilote est auto-detecte, le modele ne sert qu'aux flux
        $tplHw = $null; $clone = $false
        if ($ChkClone.IsChecked) {
            $ti = $CboTemplate.SelectedItem
            if (-not $ti -or -not $ti.Tag) {
                [System.Windows.MessageBox]::Show($script:T.AC_ValNoTemplate, $script:T.AC_Title, 'OK', 'Warning') | Out-Null
                return
            }
            $tplHw = $ti.Tag
            $clone = $true
        }

        # Confirmation : ecriture dans la configuration du VMS de production
        $confirm = [System.Windows.MessageBox]::Show(
            ($script:T.AC_Confirm -f $items.Count, "$($rsItem.Content)"),
            $script:T.AC_ConfirmTitle, 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { return }

        # Credential (SecureString, jamais journalise). PSCredential accepte un
        # SecureString vide (mot de passe vide) sans lever d'erreur.
        $cred = New-Object System.Management.Automation.PSCredential($user, $PwdBox.SecurePassword)

        $script:_AC_Result = @{
            RecServer     = $rsItem.Tag
            RecServerName = "$($rsItem.Content)"
            Items         = $items
            Credential    = $cred
            UseHttps      = [bool]$ChkHttps.IsChecked
            TemplateHw    = $tplHw
            CloneStreams  = $clone
        }
        $win.DialogResult = $true
    })

    if ($win.ShowDialog() -eq $true) { return $script:_AC_Result }
    return $null
}


function Add-CameraDevice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Config,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {}
    )

    $cfg = Show-CameraAddDialog
    if (-not $cfg) { & $Log $script:T.AC_Cancelled ; return }

    # --- Aiguillage : generation du modele Excel ou import depuis Excel ---
    if ($cfg.ExcelAction -eq 'Template') {
        & $Log $script:T.ICT_Generating
        if (New-CameraExcelTemplate -Path $cfg.Path -ServerNames $cfg.ServerNames -Log $Log) {
            & $Log ($script:T.ICT_Done -f $cfg.Path)
            try { Start-Process -FilePath 'explorer.exe' -ArgumentList "/select,`"$($cfg.Path)`"" } catch {}
        }
        return
    }
    if ($cfg.ExcelAction -eq 'Import') {
        Import-CameraExcelRows -Path $cfg.Path -Log $Log -Cancel $Cancel -ReportProgress $ReportProgress
        return
    }

    # --- Lecture optionnelle des flux du modele a cloner (pilote auto-detecte) ---
    $tplStreams = $null
    if ($cfg.CloneStreams -and $cfg.TemplateHw) {
        & $Log ($script:T.AC_LogTplLoad -f "$($cfg.TemplateHw.Name)")
        try {
            $tplCam = @($cfg.TemplateHw | Get-VmsCamera) | Select-Object -First 1
            if ($tplCam) { $tplStreams = @(Get-VmsCameraStream -Camera $tplCam) }
        } catch { & $Log ($script:T.AC_LogTplWarn -f $_.Exception.Message) }
    }

    $total = $cfg.Items.Count
    & $Log ($script:T.AC_LogStart -f $total)

    $done = 0; $ok = 0; $err = 0; $skip = 0
    foreach ($item in $cfg.Items) {
        if (& $Cancel) { break }
        $done++
        & $ReportProgress $done $total

        $addr = "$($item.Address)"
        & $Log ($script:T.AC_LogAdding -f $done, $total, $addr)

        $u = Resolve-CameraUri -Address $addr -DefaultHttps ([bool]$cfg.UseHttps)
        if (-not $u.Uri) {
            $err++
            & $Log ($script:T.AC_ValBadAddress -f $addr)
            continue
        }

        try {
            # Scan auto-detection du pilote + ajout (logique partagee avec l'import Excel)
            $res = Add-CameraByScan -RecordingServer $cfg.RecServer -Uri $u.Uri -Credential $cfg.Credential -UseHttps $u.Https
            if ($res.Status -eq 'notfound') { $err++;  & $Log ($script:T.AC_LogAddErr -f $addr, $res.Message) ; continue }
            if ($res.Status -eq 'exists')   { $skip++; & $Log ($script:T.AC_LogExists -f $addr) ; continue }
            $newHw = $res.Hardware
            $ok++
            & $Log ($script:T.AC_LogAddOk -f $addr)

            Rename-HardwareSafe -Hardware $newHw -Name "$($item.Name)" -Log $Log -LogKey 'AC_LogRenamed'

            # --- Clonage des reglages de flux ---
            if ($tplStreams -and $tplStreams.Count -gt 0) {
                try {
                    $newCams = @($newHw | Get-VmsCamera)
                    if ($newCams.Count -eq 0) {
                        & $Log ($script:T.AC_LogCloneNoCam -f $addr)
                    }
                    else {
                        $cloned = 0
                        foreach ($nc in $newCams) {
                            $ncStreams = @(Get-VmsCameraStream -Camera $nc)
                            # Lookup nom -> flux (une passe, au lieu d'un Where-Object par flux modele)
                            $byName = @{}
                            foreach ($s in $ncStreams) { $byName["$($s.Name)"] = $s }
                            for ($ti = 0; $ti -lt $tplStreams.Count; $ti++) {
                                $ts = $tplStreams[$ti]
                                if (-not $ts.Settings) { continue }
                                # Appariement par nom, repli sur l'index (noms parfois localises)
                                $match = $byName["$($ts.Name)"]
                                if (-not $match -and $ti -lt $ncStreams.Count) { $match = $ncStreams[$ti] }
                                if ($match) {
                                    Set-VmsCameraStream -Stream $match -Settings $ts.Settings -ErrorAction Stop
                                    $cloned++
                                }
                            }
                        }
                        & $Log ($script:T.AC_LogClone -f $cloned)
                    }
                }
                catch { & $Log ($script:T.AC_LogCloneErr -f $addr, $_.Exception.Message) }
            }
        }
        catch {
            $err++
            & $Log ($script:T.AC_LogAddErr -f $addr, $_.Exception.Message)
        }
    }

    if ($skip -gt 0) { & $Log ($script:T.AC_LogSkipped -f $skip) }
    & $Log ($script:T.AC_LogDone -f $ok, $err)
}
