<#
.SYNOPSIS
    Exporte un rapport Excel de tous les equipements Milestone.
    Colonnes selectionnables via une fenetre de choix. Mots de passe exclus par defaut.
#>

function Show-ExportColumnSelector {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($script:T.EH_DialogTitle)"
        Width="530" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#1E1E2E" FontFamily="Segoe UI">
    <StackPanel Margin="20,16,20,20">

        <DockPanel Margin="0,0,0,10">
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <Button x:Name="BtnSelectAll" Content="$($script:T.EH_BtnSelectAll)"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand" Margin="0,0,6,0"/>
                <Button x:Name="BtnDeselectAll" Content="$($script:T.EH_BtnDeselAll)"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand"/>
            </StackPanel>
            <TextBlock Text="$($script:T.EH_SelectCols)"
                       Foreground="#CDD6F4" FontSize="13" VerticalAlignment="Center"/>
        </DockPanel>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpHardware)"
                           Foreground="#F4D6CD" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <UniformGrid Columns="3">
                    <CheckBox x:Name="ChkNom"         Content="$($script:T.EH_ChkNom)"        IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFabricant"   Content="$($script:T.EH_ChkFabricant)"  IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkModele"      Content="$($script:T.EH_ChkModele)"     IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkIP"          Content="$($script:T.EH_ChkIP)"         IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkMAC"         Content="$($script:T.EH_ChkMAC)"        IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFirmware"    Content="$($script:T.EH_ChkFirmware)"   IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkServeurRec"  Content="$($script:T.EH_ChkServeurRec)" IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkUtilisateur" Content="$($script:T.EH_ChkUser)"       IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkMotDePasse"  Content="$($script:T.EH_ChkPassword)"   IsChecked="False" Foreground="#FAB387" FontSize="12" Margin="4,5,4,5"/>
                </UniformGrid>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpFlux)"
                           Foreground="#A8DADC" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <UniformGrid Columns="3">
                    <CheckBox x:Name="ChkCodecEnreg"      Content="$($script:T.EH_ChkCodecRec)"  IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkResolutionEnreg" Content="$($script:T.EH_ChkResRec)"    IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFPSEnreg"        Content="$($script:T.EH_ChkFpsRec)"   IsChecked="True"  Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkCodecLive"       Content="$($script:T.EH_ChkCodecLive)" IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkResolutionLive"  Content="$($script:T.EH_ChkResLive)"   IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFPSLive"         Content="$($script:T.EH_ChkFpsLive)"  IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                    <CheckBox x:Name="ChkFluxSupp"        Content="$($script:T.EH_ChkFluxSupp)" IsChecked="False" Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                </UniformGrid>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpRetention)"
                           Foreground="#A6E3A1" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <CheckBox x:Name="ChkRetention" Content="$($script:T.EH_ChkRetention)" IsChecked="True"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpOptions)"
                           Foreground="#CBA6F7" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <CheckBox x:Name="ChkImagesRef" Content="$($script:T.EH_ChkImagesRef)" IsChecked="False"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                <CheckBox x:Name="ChkSnapshotJ7" Content="$($script:T.EH_ChkSnapshotJ7)" IsChecked="True"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
                <CheckBox x:Name="ChkSnapshot" Content="$($script:T.EH_ChkSnapshot)" IsChecked="True"
                          Foreground="#CDD6F4" FontSize="12" Margin="4,5,4,5"/>
            </StackPanel>
        </Border>

        <Border Background="#181825" BorderBrush="#313244" BorderThickness="1" CornerRadius="6"
                Padding="14,10" Margin="0,0,0,8">
            <StackPanel>
                <TextBlock Text="$($script:T.EH_GrpMiseEnPage)"
                           Foreground="#F38BA8" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="$($script:T.EH_LblRowsPerPage)"
                               Foreground="#CDD6F4" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <ComboBox x:Name="CboRowsPerPage" Width="130" FontSize="12">
                        <ComboBox.Template>
                            <ControlTemplate TargetType="ComboBox">
                                <Grid>
                                    <!-- ToggleButton couvre TOUTE la largeur : clic partout possible -->
                                    <ToggleButton Focusable="False" ClickMode="Press"
                                        IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                                        <ToggleButton.Style>
                                            <Style TargetType="ToggleButton">
                                                <Setter Property="Template">
                                                    <Setter.Value>
                                                        <ControlTemplate TargetType="ToggleButton">
                                                            <Border Background="#313244" BorderBrush="#45475A"
                                                                    BorderThickness="1" CornerRadius="3">
                                                                <ContentPresenter/>
                                                            </Border>
                                                        </ControlTemplate>
                                                    </Setter.Value>
                                                </Setter>
                                            </Style>
                                        </ToggleButton.Style>
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="22"/>
                                            </Grid.ColumnDefinitions>
                                            <ContentPresenter Grid.Column="0"
                                                Content="{Binding SelectionBoxItem, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                                                ContentTemplate="{Binding SelectionBoxItemTemplate, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                                                Margin="8,4,0,4" VerticalAlignment="Center"
                                                TextBlock.Foreground="#CDD6F4" TextBlock.FontSize="12"/>
                                            <Path Grid.Column="1" Data="M 0 0 L 6 6 L 12 0"
                                                  Stroke="#CDD6F4" StrokeThickness="1.5"
                                                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Grid>
                                    </ToggleButton>
                                    <Popup x:Name="PART_Popup" Placement="Bottom"
                                           IsOpen="{TemplateBinding IsDropDownOpen}"
                                           AllowsTransparency="True" Focusable="False">
                                        <Border Background="#313244" BorderBrush="#45475A" BorderThickness="1">
                                            <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                                                <ItemsPresenter/>
                                            </ScrollViewer>
                                        </Border>
                                    </Popup>
                                </Grid>
                            </ControlTemplate>
                        </ComboBox.Template>
                        <ComboBox.ItemContainerStyle>
                            <Style TargetType="ComboBoxItem">
                                <Setter Property="Foreground" Value="#CDD6F4"/>
                                <Setter Property="Background" Value="#313244"/>
                                <Setter Property="FontSize"   Value="12"/>
                                <Setter Property="Padding"    Value="8,5"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="ComboBoxItem">
                                            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                                    Padding="{TemplateBinding Padding}">
                                                <ContentPresenter/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter TargetName="Bd" Property="Background" Value="#45475A"/>
                                                </Trigger>
                                                <Trigger Property="IsSelected" Value="True">
                                                    <Setter TargetName="Bd" Property="Background" Value="#585B70"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </ComboBox.ItemContainerStyle>
                        <ComboBoxItem Content="$($script:T.EH_RppUnlimited)" IsSelected="True"/>
                        <ComboBoxItem Content="5"/>
                        <ComboBoxItem Content="10"/>
                        <ComboBoxItem Content="15"/>
                        <ComboBoxItem Content="20"/>
                        <ComboBoxItem Content="25"/>
                        <ComboBoxItem Content="50"/>
                    </ComboBox>
                </StackPanel>
            </StackPanel>
        </Border>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="BtnCancel" Content="$($script:T.EH_BtnCancel)" IsCancel="True"
                    Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                    FontSize="13" Padding="22,8" Cursor="Hand" Margin="0,0,10,0"/>
            <Button x:Name="BtnExport" Content="$($script:T.EH_BtnExport)" IsDefault="True"
                    Background="#A6E3A1" Foreground="#1E1E2E" BorderThickness="0"
                    FontWeight="Bold" FontSize="13" Padding="22,8" Cursor="Hand"/>
        </StackPanel>

    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $checkboxes = [ordered]@{
        'Nom'             = $window.FindName('ChkNom')
        'Fabricant'       = $window.FindName('ChkFabricant')
        'Modele'          = $window.FindName('ChkModele')
        'IP'              = $window.FindName('ChkIP')
        'MAC'             = $window.FindName('ChkMAC')
        'Firmware'        = $window.FindName('ChkFirmware')
        'ServeurRec'      = $window.FindName('ChkServeurRec')
        'Utilisateur'     = $window.FindName('ChkUtilisateur')
        'MotDePasse'      = $window.FindName('ChkMotDePasse')
        'CodecEnreg'      = $window.FindName('ChkCodecEnreg')
        'ResolutionEnreg' = $window.FindName('ChkResolutionEnreg')
        'FPSEnreg'        = $window.FindName('ChkFPSEnreg')
        'CodecLive'       = $window.FindName('ChkCodecLive')
        'ResolutionLive'  = $window.FindName('ChkResolutionLive')
        'FPSLive'         = $window.FindName('ChkFPSLive')
        'FluxSupp'        = $window.FindName('ChkFluxSupp')
        'Retention'       = $window.FindName('ChkRetention')
        'ImagesRef'       = $window.FindName('ChkImagesRef')
        'SnapshotJ7'      = $window.FindName('ChkSnapshotJ7')
        'Snapshot'        = $window.FindName('ChkSnapshot')
    }

    # Demande a l'utilisateur de choisir le dossier des images de reference.
    # Decoche la case si l'utilisateur annule la selection.
    function Select-RefImagesFolder {
        param([System.Windows.Controls.CheckBox]$Checkbox)
        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg.Description       = $script:T.EH_RefImagesFolderPrompt
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $Checkbox.Tag = $dlg.SelectedPath
            return $true
        }
        return $false
    }

    $checkboxes['ImagesRef'].Add_Click({
        $chk = $checkboxes['ImagesRef']
        if ($chk.IsChecked -eq $true -and -not (Select-RefImagesFolder -Checkbox $chk)) {
            $chk.IsChecked = $false
        }
    })

    $window.FindName('BtnSelectAll').Add_Click({
        foreach ($chk in $checkboxes.Values) { $chk.IsChecked = $true }
    })

    $window.FindName('BtnDeselectAll').Add_Click({
        foreach ($chk in $checkboxes.Values) { $chk.IsChecked = $false }
    })

    $window.FindName('BtnExport').Add_Click({
        # Filet de securite : si la case est cochee via "Tout cocher" sans
        # passer par son gestionnaire de clic, le dossier n'a pas ete choisi.
        $chkRef = $checkboxes['ImagesRef']
        if ($chkRef.IsChecked -eq $true -and [string]::IsNullOrWhiteSpace([string]$chkRef.Tag)) {
            if (-not (Select-RefImagesFolder -Checkbox $chkRef)) { $chkRef.IsChecked = $false }
        }

        $sel = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $checkboxes.Keys) {
            if ($checkboxes[$name].IsChecked -eq $true) { $sel.Add($name) }
        }
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                $script:T.EH_NoColumn, $script:T.EH_NoColumnTitle,
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
            return
        }
        $rppRaw = $window.FindName('CboRowsPerPage').SelectedItem.Content
        $rpp    = if ($rppRaw -eq $script:T.EH_RppUnlimited) { 0 } else { [int]$rppRaw }
        $window.Tag = @{ Columns = [string[]]$sel; RowsPerPage = $rpp; RefImagesFolder = [string]$chkRef.Tag }
        $window.DialogResult = $true
    })

    if ($window.ShowDialog() -eq $true) { return $window.Tag }
    return $null
}


function Export-HardwareReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Config,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {}
    )

    function Get-StreamSetting {
        param($stream, [string]$key)
        if ($stream -and $stream.Settings -and $stream.Settings[$key]) { return $stream.Settings[$key] }
        return 'N/A'
    }

    # Valeurs des colonnes d'une camera — SOURCE UNIQUE partagee entre le chemin
    # ImportExcel et le chemin COM (evite toute divergence entre les deux exports).
    function Get-CameraRowValues {
        param($cam, [hashtable]$StreamLookup, [hashtable]$RetentionLookup, [bool]$IncludePassword)
        $ip = if ($cam.Address -match '([0-9]{1,3}(?:\.[0-9]{1,3}){3})') { $Matches[1] } else { $cam.Address }
        $si         = $StreamLookup[$cam.Name]
        $recStream  = if ($si) { $si.Rec  } else { $null }
        $liveStream = if ($si) { $si.Live } else { $null }
        $extraCount = if ($si) { $si.Extra } else { 0 }
        # Meme flux physique pour enregistrement et live => on n'affiche pas deux fois
        # les memes valeurs (les colonnes live restent vides).
        $sameStream = $recStream -and $liveStream -and
                      ($recStream.StreamReferenceId -eq $liveStream.StreamReferenceId)
        $ret        = if ($RetentionLookup.ContainsKey($cam.Name)) { $RetentionLookup[$cam.Name] } else { 'N/A' }
        @{
            'Nom'             = $cam.Name
            'Fabricant'       = $cam.DriverFamily
            'Modele'          = $cam.Model
            'IP'              = $ip
            'MAC'             = $cam.MAC
            'Firmware'        = $cam.Firmware
            'ServeurRec'      = $cam.RecorderName
            'Utilisateur'     = $cam.Username
            'MotDePasse'      = if ($IncludePassword) { $cam.Password } else { '' }
            'CodecEnreg'      = Get-StreamSetting $recStream 'Codec'
            'ResolutionEnreg' = Get-StreamSetting $recStream 'Resolution'
            'FPSEnreg'        = Get-StreamSetting $recStream 'FPS'
            'CodecLive'       = if ($sameStream) { '' } else { Get-StreamSetting $liveStream 'Codec' }
            'ResolutionLive'  = if ($sameStream) { '' } else { Get-StreamSetting $liveStream 'Resolution' }
            'FPSLive'         = if ($sameStream) { '' } else { Get-StreamSetting $liveStream 'FPS' }
            'FluxSupp'        = if ($extraCount -gt 0) { $script:T.XL_ExtraFlux -f $extraCount } else { '' }
            'Retention'       = $ret
        }
    }

    $exportConfig = Show-ExportColumnSelector
    if ($null -eq $exportConfig) {
        & $Log $script:T.EH_Cancelled
        return
    }
    $selectedColumns = $exportConfig.Columns
    $rowsPerPage     = [int]$exportConfig.RowsPerPage

    $includePassword    = $selectedColumns -contains 'MotDePasse'
    $includeSnapshots   = $selectedColumns -contains 'Snapshot'
    $includeSnapshotsJ7 = $selectedColumns -contains 'SnapshotJ7'
    $includeRefImages   = $selectedColumns -contains 'ImagesRef'
    $refImagesFolder    = $exportConfig.RefImagesFolder
    $needStreams       = ($selectedColumns | Where-Object {
        $_ -in 'CodecEnreg','ResolutionEnreg','FPSEnreg','CodecLive','ResolutionLive','FPSLive','FluxSupp'
    }).Count -gt 0
    $needRetention    = $selectedColumns -contains 'Retention'

    # Definition des colonnes dans l'ordre canonique
    $allColumnDefs = @(
        @{ Name = 'Nom';             Group = 'base';   Header = $script:T.XL_Nom }
        @{ Name = 'Fabricant';       Group = 'base';   Header = $script:T.XL_Fabricant }
        @{ Name = 'Modele';          Group = 'base';   Header = $script:T.XL_Modele }
        @{ Name = 'IP';              Group = 'base';   Header = $script:T.XL_IP }
        @{ Name = 'MAC';             Group = 'base';   Header = $script:T.XL_MAC }
        @{ Name = 'Firmware';        Group = 'base';   Header = $script:T.XL_Firmware }
        @{ Name = 'ServeurRec';      Group = 'base';   Header = $script:T.XL_ServeurRec }
        @{ Name = 'Utilisateur';     Group = 'base';   Header = $script:T.XL_Utilisateur }
        @{ Name = 'MotDePasse';      Group = 'base';   Header = $script:T.XL_MotDePasse }
        @{ Name = 'CodecEnreg';      Group = 'stream'; Header = $script:T.XL_CodecEnreg }
        @{ Name = 'ResolutionEnreg'; Group = 'stream'; Header = $script:T.XL_ResEnreg }
        @{ Name = 'FPSEnreg';        Group = 'stream'; Header = $script:T.XL_FpsEnreg }
        @{ Name = 'CodecLive';       Group = 'stream'; Header = $script:T.XL_CodecLive }
        @{ Name = 'ResolutionLive';  Group = 'stream'; Header = $script:T.XL_ResLive }
        @{ Name = 'FPSLive';         Group = 'stream'; Header = $script:T.XL_FpsLive }
        @{ Name = 'FluxSupp';        Group = 'stream'; Header = $script:T.XL_FluxSupp }
        @{ Name = 'Retention';    Group = 'ret';  Header = $script:T.XL_Retention }
        @{ Name = 'ImagesRef';   Group = 'snap'; Header = $script:T.XL_ImagesRef }
        @{ Name = 'SnapshotJ7';  Group = 'snap'; Header = $script:T.XL_SnapshotJ7 }
        @{ Name = 'Snapshot';    Group = 'snap'; Header = $script:T.XL_Snapshot }
    )

    $activeColumns = @($allColumnDefs | Where-Object { $selectedColumns -contains $_.Name })

    $snapColIndex   = 0
    $snapJ7ColIndex = 0
    $refImgColIndex = 0
    for ($i = 0; $i -lt $activeColumns.Count; $i++) {
        if ($activeColumns[$i].Name -eq 'Snapshot')   { $snapColIndex   = $i + 1 }
        if ($activeColumns[$i].Name -eq 'SnapshotJ7') { $snapJ7ColIndex = $i + 1 }
        if ($activeColumns[$i].Name -eq 'ImagesRef')  { $refImgColIndex = $i + 1 }
    }

    & $Log $script:T.EH_LogGenerating
    if ($includePassword) {
        $camReport = @(Get-VmsCameraReport -IncludePlainTextPassword)
    } else {
        $camReport = @(Get-VmsCameraReport)
    }
    $total = $camReport.Count
    & $Log ($script:T.EH_LogFound -f $total)

    & $Log $script:T.EH_LogLoadCams
    $vmsCameras   = @(Get-VmsCamera)
    $vmsCamByName = @{}
    $vmsCamByPath = @{}
    foreach ($c in $vmsCameras) {
        $vmsCamByName[$c.Name] = $c
        $vmsCamByPath[$c.Path] = $c.Name
    }

    $streamLookup = @{}
    if ($needStreams) {
        & $Log $script:T.EH_LogStreams
        try {
            $allStreams = @($vmsCameras | Get-VmsCameraStream -Enabled -ErrorAction Stop)

            # Regroupe d'abord TOUS les flux par camera, puis identifie enregistre/live
            # de maniere INDEPENDANTE via les drapeaux de configuration reels — un meme
            # flux peut etre a la fois enregistre ET live par defaut (cas le plus courant).
            # L'ancien if/elseif classait un tel flux uniquement en "enregistre" et
            # laissait le live vide.
            $streamsByCam = @{}
            foreach ($s in $allStreams) {
                $name = $s.Camera.Name
                if (-not $streamsByCam.ContainsKey($name)) {
                    $streamsByCam[$name] = [System.Collections.Generic.List[object]]::new()
                }
                $streamsByCam[$name].Add($s)
            }

            foreach ($name in $streamsByCam.Keys) {
                $streams = $streamsByCam[$name]

                # Flux enregistre : on ne considere QUE les flux dont le booleen Recorded
                # est vrai (signal sans ambiguite). En multi-track, on prefere la piste
                # primaire ; le nom de piste est expose selon les versions sur
                # RecordingTrackName ou RecordingTrack, on teste les deux.
                $recorded = @($streams | Where-Object { $_.Recorded })
                $rec = $recorded | Where-Object {
                    $_.RecordingTrackName -eq 'Primary' -or $_.RecordingTrack -eq 'Primary'
                } | Select-Object -First 1
                if (-not $rec) { $rec = $recorded | Select-Object -First 1 }

                # Flux live : celui marque "LiveDefault" (peut etre le meme objet que $rec)
                $live = $streams | Where-Object { $_.LiveDefault } | Select-Object -First 1

                # Flux supplementaires : tous ceux qui ne sont ni l'enregistre ni le live retenus
                $keepIds = @()
                if ($rec)  { $keepIds += $rec.StreamReferenceId }
                if ($live) { $keepIds += $live.StreamReferenceId }
                $extra = @($streams | Where-Object { $_.StreamReferenceId -notin $keepIds }).Count

                $streamLookup[$name] = @{ Rec = $rec; Live = $live; Extra = $extra }
            }

            & $Log ($script:T.EH_LogStreamsOk -f $allStreams.Count, $streamLookup.Count)
        }
        catch { & $Log ($script:T.EH_LogStreamsErr -f $_) }
    }

    $retentionLookup = @{}
    if ($needRetention) {
        & $Log $script:T.EH_LogPlayback
        try {
            $playbackData = @($vmsCameras | Get-PlaybackInfo -Parallel -ErrorAction Stop)
            foreach ($pb in $playbackData) {
                $name = $vmsCamByPath[$pb.Path]
                if ($name) {
                    if ($pb.Begin -and $pb.End) {
                        $days = [int]($pb.End - $pb.Begin).TotalDays
                        $retentionLookup[$name] = "$days j"
                    }
                    else { $retentionLookup[$name] = $script:T.XL_Aucun }
                }
            }
            & $Log ($script:T.EH_LogPlaybackOk -f $retentionLookup.Count)
        }
        catch { & $Log ($script:T.EH_LogPlaybackErr -f $_) }
    }

    $tempDir    = $null
    $tempDirJ7  = $null
    $snapPaths  = @{}
    $snapJ7Paths = @{}

    if ($includeSnapshots) {
        & $Log $script:T.EH_LogSnaps
        $quality  = $Config.snapshotQuality
        $tempDir  = Join-Path $env:TEMP "MilestoneHW_$(Get-Random)"
        New-Item $tempDir -ItemType Directory -Force | Out-Null

        $received = 0
        $snapTotal = $camReport.Count

        foreach ($cam in $camReport) {
            if (& $Cancel) { break }
            $vmsCamera = $vmsCamByName[$cam.Name]
            if (-not $vmsCamera) { continue }

            $safeName = $cam.Name -replace '[\\/:*?"<>|]', '_'
            $filePath = Join-Path $tempDir "$safeName.jpg"

            try {
                $snap = $vmsCamera | Get-Snapshot -Behavior GetEnd -Quality $quality -ErrorAction Stop
                if ($snap -and $snap.Bytes -and $snap.Bytes.Length -gt 0) {
                    [System.IO.File]::WriteAllBytes($filePath, $snap.Bytes)
                    $snapPaths[$cam.Name] = $filePath
                    $received++
                    & $Log ($script:T.EH_LogSnapOk -f $received, $snapTotal, $cam.Name)
                } else {
                    & $Log ($script:T.EH_LogSnapEmpty -f $cam.Name)
                }
            } catch {
                & $Log ($script:T.EH_LogSnapErr -f $cam.Name, $_)
            }

            & $ReportProgress $received $snapTotal
        }

        & $Log ($script:T.EH_LogSnapsDone -f $snapPaths.Count, $snapTotal)
    }

    if ($includeSnapshotsJ7) {
        $j7Time   = [datetime]::Now.AddDays(-7)
        $quality  = $Config.snapshotQuality
        & $Log ($script:T.EH_LogSnapsJ7 -f $j7Time.ToString('dd/MM/yyyy HH:mm'))
        $tempDirJ7 = Join-Path $env:TEMP "MilestoneHW_J7_$(Get-Random)"
        New-Item $tempDirJ7 -ItemType Directory -Force | Out-Null

        $j7Total = $camReport.Count ; $j7Recv = 0
        foreach ($cam in $camReport) {
            if (& $Cancel) { break }
            $vmsCamera = $vmsCamByName[$cam.Name]
            if (-not $vmsCamera) { continue }
            $safeName = $cam.Name -replace '[\\/:*?"<>|]', '_'
            $filePath = Join-Path $tempDirJ7 "$safeName.jpg"
            try {
                $snap = $vmsCamera | Get-Snapshot -Behavior GetNearest -Time $j7Time -Quality $quality -ErrorAction Stop
                if ($snap -and $snap.Bytes -and $snap.Bytes.Length -gt 0) {
                    [System.IO.File]::WriteAllBytes($filePath, $snap.Bytes)
                    $snapJ7Paths[$cam.Name] = $filePath
                    $j7Recv++
                    & $Log ($script:T.EH_LogSnapOk -f $j7Recv, $j7Total, $cam.Name)
                } else { & $Log ($script:T.EH_LogSnapEmpty -f $cam.Name) }
            } catch { & $Log ($script:T.EH_LogSnapErr -f $cam.Name, $_) }
        }
        & $Log ($script:T.EH_LogSnapsDone -f $snapJ7Paths.Count, $j7Total)
    }

    $refImgPaths = @{}
    if ($includeRefImages -and $refImagesFolder -and (Test-Path $refImagesFolder)) {
        & $Log ($script:T.EH_LogRefImages -f $refImagesFolder)
        $refExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.gif')
        # Tri une seule fois (du plus recent au plus ancien) : la recherche par camera
        # peut alors se contenter d'un "Select -First 1" sans re-trier a chaque iteration.
        $refFiles = @(Get-ChildItem -Path $refImagesFolder -File -ErrorAction SilentlyContinue |
            Where-Object { $refExtensions -contains $_.Extension.ToLowerInvariant() } |
            Sort-Object LastWriteTime -Descending)
        $refTotal = $camReport.Count
        foreach ($cam in $camReport) {
            if (& $Cancel) { break }
            $safeName = $cam.Name -replace '[\\/:*?"<>|]', '_'
            # Correspond a "NomCamera.ext" ou "NomCamera_<suffixe>.ext" (ex. snapshots horodates)
            $match = $refFiles | Where-Object {
                $_.BaseName -eq $safeName -or $_.BaseName.StartsWith("${safeName}_", [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1
            if ($match) {
                $refImgPaths[$cam.Name] = $match.FullName
                & $Log ($script:T.EH_LogSnapOk -f $refImgPaths.Count, $refTotal, $cam.Name)
            } else {
                & $Log ($script:T.EH_LogRefImgMissing -f $cam.Name)
            }
        }
        & $Log ($script:T.EH_LogRefImagesDone -f $refImgPaths.Count, $refTotal)
    }

    if (-not (Test-Path $Config.outputDirectory)) {
        New-Item -Path $Config.outputDirectory -ItemType Directory -Force | Out-Null
    }
    $xlsxPath = [System.IO.Path]::GetFullPath((Join-Path $Config.outputDirectory $script:T.XL_FileName))

    # Couleurs des groupes de colonnes — partagees entre le chemin COM et ImportExcel
    # Format COM : BGR (0xBBGGRR) — extrait en RGB pour System.Drawing / EPPlus
    $groupColors = @{
        'base'   = @{ Bg = 0x44413D; Fg = 0xF4D6CD }
        'stream' = @{ Bg = 0x1D3557; Fg = 0xA8DADC }
        'ret'    = @{ Bg = 0x1B4332; Fg = 0xA6E3A1 }
        'snap'   = @{ Bg = 0x2D2B55; Fg = 0xCBA6F7 }
    }

    function Get-ImportExcelModule {
        # Module already imported in this session (e.g. loaded from Dependencies/ at startup)
        if (Get-Module -Name ImportExcel) {
            return $true
        }

        if (Get-Module -ListAvailable -Name ImportExcel) {
            try {
                Import-Module ImportExcel -Force -ErrorAction Stop
                return $true
            }
            catch {}
        }

        # Mode Offline (dossier Dependencies/ present) : ne jamais telecharger
        # depuis Internet en pleine action — le module doit venir du cache local.
        if ($script:DependenciesPath -and (Test-Path $script:DependenciesPath)) {
            return $false
        }

        # Annonce le telechargement au lieu d'installer silencieusement
        & $Log $script:T.EH_LogInstallExcel
        try {
            Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
            Import-Module ImportExcel -Force -ErrorAction Stop
            return $true
        }
        catch {
            return $false
        }
    }

    # Tenter Excel COM — echec silencieux, on bascule sur ImportExcel si absent
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application -ErrorAction Stop }
    catch { & $Log $script:T.EH_LogNoExcel }

    # ---------------------------------------------------------------
    # CHEMIN IMPORTEXCEL — hors du catch pour eviter les anomalies de
    # propagation d'erreur de PowerShell 5.1 dans un bloc catch WPF.
    # ---------------------------------------------------------------
    if (-not $excel) {
        if (-not (Get-ImportExcelModule)) {
            & $Log $script:T.EH_LogNoExcelNoData
            if ($tempDir -and (Test-Path $tempDir)) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
            return
        }
        & $Log $script:T.EH_LogExcelFallback

        if ($activeColumns.Count -eq 0) {
            & $Log $script:T.EH_LogNoExcelNoData
            if ($tempDir -and (Test-Path $tempDir)) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
            return
        }

        # ---------------------------------------------------------------
        # CHEMIN IMPORTEXCEL : tout est delegue au subprocess (EPPlus propre).
        # Export-Excel depuis le contexte WPF+MilestonePSTools provoque des
        # erreurs EPPlus internes. Le subprocess cree le xlsx de zero.
        # ---------------------------------------------------------------
        & $Log $script:T.EH_LogSubproc

        # Calculer les valeurs de chaque camera (stream/retention deja en memoire)
        $camValues = [System.Collections.Generic.List[pscustomobject]]::new()
        foreach ($cam in $camReport) {
            $epRow  = [ordered]@{}
            $values = Get-CameraRowValues $cam $streamLookup $retentionLookup $includePassword
            foreach ($col in $activeColumns) {
                $v = $values[$col.Name]
                $epRow[$col.Name] = if ($null -ne $v) { $v } else { '' }
            }
            $epRow['__Snap']   = if ($snapColIndex   -gt 0 -and $snapPaths.ContainsKey($cam.Name))   { $snapPaths[$cam.Name]   } else { '' }
            $epRow['__SnapJ7'] = if ($snapJ7ColIndex -gt 0 -and $snapJ7Paths.ContainsKey($cam.Name)) { $snapJ7Paths[$cam.Name] } else { '' }
            $epRow['__RefImg'] = if ($refImgColIndex -gt 0 -and $refImgPaths.ContainsKey($cam.Name)) { $refImgPaths[$cam.Name] } else { '' }
            $camValues.Add([pscustomobject]$epRow)
        }

        # Tout le reste (creation xlsx, couleurs, bordures, snapshots) est fait
        # dans un subprocess PowerShell propre — aucun appel EPPlus dans ce process.
        $spPayload = [ordered]@{
            XlsxPath    = $xlsxPath
            ImExPath    = (Get-Module ImportExcel).ModuleBase
            SheetBase   = $script:T.XL_SheetName
            RowsPerPage = $rowsPerPage
            SnapCol     = $snapColIndex
            SnapJ7Col   = $snapJ7ColIndex
            RefImgCol   = $refImgColIndex
            GroupColors = @{
                base   = @{ Bg = $groupColors.base.Bg;   Fg = $groupColors.base.Fg }
                stream = @{ Bg = $groupColors.stream.Bg; Fg = $groupColors.stream.Fg }
                ret    = @{ Bg = $groupColors.ret.Bg;    Fg = $groupColors.ret.Fg }
                snap   = @{ Bg = $groupColors.snap.Bg;   Fg = $groupColors.snap.Fg }
            }
            Columns     = @($activeColumns | ForEach-Object { [ordered]@{ Name=$_.Name; Group=$_.Group; Header=$_.Header } })
            Cameras     = $camValues
        } | ConvertTo-Json -Depth 5

        $spJson = Join-Path $env:TEMP "MHW_$(Get-Random).json"
        $spPs1  = Join-Path $env:TEMP "MHW_$(Get-Random).ps1"
        Set-Content -Path $spJson -Value $spPayload -Encoding UTF8

        $spScript = @'
param([string]$J)
try {
    $d = Get-Content $J -Raw | ConvertFrom-Json
    Import-Module $d.ImExPath -Force
    Add-Type -AssemblyName System.Drawing
    function cc([int]$n){ $s='';$t=$n;while($t-gt 0){$t--;$s=[char]([byte][char]'A'+($t%26))+$s;$t=[Math]::Floor($t/26)};$s }
    $light  = [System.Drawing.Color]::FromArgb(221,235,247)
    $dark   = [System.Drawing.Color]::FromArgb(189,215,238)
    $bs     = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $nbCols = $d.Columns.Count
    $lastCol = cc $nbCols
    $total  = $d.Cameras.Count
    $rpp    = [int]$d.RowsPerPage
    $chunk  = if ($rpp -gt 0) { $rpp } else { [Math]::Max(1,$total) }
    $pkg    = $null
    $start  = 0

    while ($start -lt $total) {
        $end   = [Math]::Min($start + $chunk, $total)
        $cams  = @($d.Cameras[$start..($end-1)])
        $sname = if ($rpp -gt 0) { "$($d.SheetBase) $($start+1)-$end" } else { $d.SheetBase }

        $rowList = [System.Collections.Generic.List[pscustomobject]]::new()
        foreach ($cam in $cams) {
            $row = [ordered]@{}
            foreach ($col in $d.Columns) {
                $row[$col.Header] = if ($col.Name -notin @('Snapshot','SnapshotJ7','ImagesRef')) { [string]$cam.($col.Name) } else { '' }
            }
            $rowList.Add([pscustomobject]$row)
        }
        $rows = $rowList.ToArray()

        if ($null -eq $pkg) {
            $pkg = $rows | Export-Excel -Path $d.XlsxPath -WorksheetName $sname -AutoSize -AutoFilter -BoldTopRow -PassThru
        } else {
            $pkg = $rows | Export-Excel -ExcelPackage $pkg -WorksheetName $sname -AutoSize -AutoFilter -BoldTopRow -PassThru
        }
        $ws = $pkg.Workbook.Worksheets[$pkg.Workbook.Worksheets.Count]
        if (-not $ws) { $start += $chunk; continue }

        $nbR = $end - $start
        $lrow = $nbR + 1

        # Couleurs en-tetes
        for ($ci = 0; $ci -lt $d.Columns.Count; $ci++) {
            $grp = $d.GroupColors.($d.Columns[$ci].Group)
            $bgH=[int]$grp.Bg; $fgH=[int]$grp.Fg
            $bgC=[System.Drawing.Color]::FromArgb($bgH -band 0xFF,($bgH -shr 8) -band 0xFF,($bgH -shr 16) -band 0xFF)
            $fgC=[System.Drawing.Color]::FromArgb($fgH -band 0xFF,($fgH -shr 8) -band 0xFF,($fgH -shr 16) -band 0xFF)
            $a = "$(cc ($ci+1))1"
            try { Set-ExcelRange -Worksheet $ws -Range $ws.Cells[$a] -BackgroundColor $bgC -FontColor $fgC } catch {}
        }

        # Lignes alternees
        for ($r = 2; $r -le $lrow; $r++) {
            $bg = if ($r % 2 -eq 0) { $light } else { $dark }
            $ws.Cells["A${r}:${lastCol}${r}"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $ws.Cells["A${r}:${lastCol}${r}"].Style.Fill.BackgroundColor.SetColor($bg)
        }

        # Bordures
        $ws.Cells["A1:${lastCol}${lrow}"].Style.Border.Top.Style    = $bs
        $ws.Cells["A1:${lastCol}${lrow}"].Style.Border.Bottom.Style = $bs
        $ws.Cells["A1:${lastCol}${lrow}"].Style.Border.Left.Style   = $bs
        $ws.Cells["A1:${lastCol}${lrow}"].Style.Border.Right.Style  = $bs

        # Snapshots — largeur colonne en characterWidth EPPlus (≈ 7px/char a 96dpi)
        # 194px / 7 ≈ 27.7 → 28. Hauteur en points : 90pt * 1.333 ≈ 120px > 118px image.
        foreach ($sd in @(
            @{ Col=[int]$d.RefImgCol; Key='__RefImg' }
            @{ Col=[int]$d.SnapJ7Col; Key='__SnapJ7' }
            @{ Col=[int]$d.SnapCol;   Key='__Snap'   }
        )) {
            if ($sd.Col -gt 0) {
                $ws.Column($sd.Col).Width = 28
                $ec = $sd.Col - 1
                for ($i = 0; $i -lt $cams.Count; $i++) {
                    $wsRow = $i + 2
                    if ($ws.Row($wsRow).Height -lt 90) { $ws.Row($wsRow).Height = 90 }
                    $sf = [string]($cams[$i].($sd.Key))
                    if ($sf -and (Test-Path $sf)) {
                        try {
                            $img = [System.Drawing.Image]::FromFile($sf)
                            $pic = $ws.Drawings.AddPicture([guid]::NewGuid().ToString(),$img)
                            $pic.SetPosition($wsRow-1, 2, $ec, 2)
                            $pic.SetSize(190, 114)
                            $img.Dispose()
                        } catch {}
                    }
                }
            }
        }
        $start += $chunk
    }
    if ($null -ne $pkg) { $pkg.Save(); $pkg.Dispose() }
} finally { Remove-Item $J -Force -ErrorAction SilentlyContinue }
'@
        Set-Content -Path $spPs1 -Value $spScript -Encoding UTF8

        if (Test-Path $xlsxPath) {
            try   { Remove-Item $xlsxPath -Force -ErrorAction Stop }
            catch { & $Log $script:T.EH_LogExcelLocked }
        }

        try {
            $proc = Start-Process -FilePath 'powershell.exe' `
                -ArgumentList '-ExecutionPolicy','Bypass','-NonInteractive','-File',$spPs1,$spJson `
                -Wait -PassThru -WindowStyle Hidden
            if ($proc.ExitCode -eq 0) { & $Log ($script:T.EH_LogSaved -f $xlsxPath) }
            else { & $Log ("AVERTISSEMENT: Sous-processus Excel code $($proc.ExitCode).") }
        } catch {
            & $Log ("AVERTISSEMENT: Sous-processus Excel : $_")
        } finally {
            Remove-Item $spPs1 -Force -ErrorAction SilentlyContinue
        }


        if ($tempDir   -and (Test-Path $tempDir))   { Remove-Item $tempDir   -Recurse -Force -ErrorAction SilentlyContinue }
        if ($tempDirJ7 -and (Test-Path $tempDirJ7)) { Remove-Item $tempDirJ7 -Recurse -Force -ErrorAction SilentlyContinue }
        return
    }

    # ---------------------------------------------------------------
    # CHEMIN COM EXCEL
    # ---------------------------------------------------------------
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    # Scriptblock : applique les en-tetes colores sur une feuille COM
    $addComHeaders = {
        param([object]$sh)
        for ($c = 0; $c -lt $activeColumns.Count; $c++) {
            $col  = $activeColumns[$c]
            $cell = $sh.Cells.Item(1, $c + 1)
            $cell.Value2              = $col.Header
            $cell.Font.Bold           = $true
            $cell.Font.Size           = 11
            $cell.HorizontalAlignment = -4108
            $cell.Interior.Color      = $groupColors[$col.Group].Bg
            $cell.Font.Color          = $groupColors[$col.Group].Fg
        }
        $sh.Rows.Item(2).Select() | Out-Null
        $sh.Application.ActiveWindow.FreezePanes = $true
    }

    # Scriptblock : finalise une feuille COM (largeurs, bordures)
    $finishComSheet = {
        param([object]$sh, [int]$lastRow)
        if ($snapColIndex   -gt 0) { $sh.Columns.Item($snapColIndex).ColumnWidth   = 28 }
        if ($snapJ7ColIndex -gt 0) { $sh.Columns.Item($snapJ7ColIndex).ColumnWidth = 28 }
        if ($refImgColIndex -gt 0) { $sh.Columns.Item($refImgColIndex).ColumnWidth = 28 }
        for ($c = 1; $c -le $activeColumns.Count; $c++) {
            if ($c -ne $snapColIndex -and $c -ne $snapJ7ColIndex -and $c -ne $refImgColIndex) { $sh.Columns.Item($c).AutoFit() | Out-Null }
        }
        if ($lastRow -ge 2) {
            $rng = $sh.Range($sh.Cells.Item(1, 1), $sh.Cells.Item($lastRow, $activeColumns.Count))
            $rng.Borders.LineStyle = 1
            $rng.Borders.Weight    = 2
        }
    }

    # Calcule le nom de feuille avec la plage de cameras : "Cameras 1-5"
    $getSheetName = {
        param([int]$from, [int]$rpp, [int]$totalCams)
        if ($rpp -le 0) { return $script:T.XL_SheetName }
        $camFrom = ($from - 1) * $rpp + 1
        $camTo   = [Math]::Min($from * $rpp, $totalCams)
        return "$($script:T.XL_SheetName) $camFrom-$camTo"
    }

    try {
        $workbook   = $excel.Workbooks.Add()
        $sheetNum   = 1
        $camOnSheet = 0
        $sheet      = $workbook.Sheets.Item(1)
        $sheet.Name = & $getSheetName 1 $rowsPerPage $total
        & $addComHeaders $sheet

        $row   = 2
        $count = 0
        & $Log $script:T.EH_LogBuilding

        foreach ($cam in $camReport) {
            if (& $Cancel) {
                & $Log ($script:T.EH_LogCancelled -f $count, $total)
                break
            }

            # Nouvelle feuille si le quota de cameras est atteint
            if ($rowsPerPage -gt 0 -and $camOnSheet -ge $rowsPerPage) {
                & $finishComSheet $sheet ($row - 1)
                $sheetNum++
                $newSheet = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
                $newSheet.Name = & $getSheetName $sheetNum $rowsPerPage $total
                $sheet = $newSheet
                & $addComHeaders $sheet
                $row        = 2
                $camOnSheet = 0
            }

            $count++
            $camOnSheet++
            & $ReportProgress $count $total
            & $Log ($script:T.EH_LogCamRow -f $count, $total, $cam.Name)

            $values = Get-CameraRowValues $cam $streamLookup $retentionLookup $includePassword

            for ($c = 0; $c -lt $activeColumns.Count; $c++) {
                $colName = $activeColumns[$c].Name
                if ($colName -in 'Snapshot','SnapshotJ7','ImagesRef') { continue }
                $sheet.Cells.Item($row, $c + 1) = $values[$colName]
            }

            # Lignes alternees (couleurs resetees a chaque feuille car $row repart de 2)
            $rowBg = if ($row % 2 -eq 0) { 0xF7EBDD } else { 0xEED7BD }
            $sheet.Range(
                $sheet.Cells.Item($row, 1),
                $sheet.Cells.Item($row, $activeColumns.Count)
            ).Interior.Color = $rowBg

            # Images : reference, puis J-7, puis Live
            if ($snapColIndex -gt 0 -or $snapJ7ColIndex -gt 0 -or $refImgColIndex -gt 0) {
                $sheet.Rows.Item($row).RowHeight = 90
            }
            foreach ($snapDef in @(
                @{ Idx = $refImgColIndex; Paths = $refImgPaths }
                @{ Idx = $snapJ7ColIndex; Paths = $snapJ7Paths }
                @{ Idx = $snapColIndex;   Paths = $snapPaths   }
            )) {
                if ($snapDef.Idx -gt 0) {
                    $sf = $snapDef.Paths[$cam.Name]
                    if ($sf -and (Test-Path $sf)) {
                        try {
                            $cell  = $sheet.Cells.Item($row, $snapDef.Idx)
                            $shape = $sheet.Shapes.AddPicture(
                                $sf, 0, -1,
                                [double]$cell.Left, [double]$cell.Top,
                                [double]$cell.Width, [double]$cell.Height
                            )
                            $shape.Placement = 1
                        } catch { & $Log ($script:T.EH_LogImgErr -f $cam.Name, $_) }
                    }
                }
            }

            $row++
        }

        # Finaliser la derniere feuille
        & $finishComSheet $sheet ($row - 1)

        $workbook.SaveAs($xlsxPath, 51)
        & $Log ($script:T.EH_LogSaved -f $xlsxPath)
    }
    finally {
        try { $workbook.Close($false) } catch {}
        try { $excel.Quit() }           catch {}
        # Liberer TOUS les objets COM references (sheet + workbook + excel), puis
        # forcer le GC : sinon les RCW restants gardent un Excel.exe zombie en memoire.
        foreach ($com in @($sheet, $workbook, $excel)) {
            if ($com) {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($com) | Out-Null } catch {}
            }
        }
        $sheet = $null ; $workbook = $null ; $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        if ($tempDirJ7 -and (Test-Path $tempDirJ7)) {
            Remove-Item $tempDirJ7 -Recurse -Force -ErrorAction SilentlyContinue
        }
        if ($tempDir -and (Test-Path $tempDir)) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
