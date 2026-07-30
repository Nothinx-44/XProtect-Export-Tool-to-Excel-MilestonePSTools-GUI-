<#
.SYNOPSIS
    Creation d'alarmes Milestone en masse, ou par duplication d'une alarme existante,
    et GESTION en masse (activer / desactiver / supprimer / nettoyer) des alarmes.
.DESCRIPTION
    Ouvre une fenetre avec deux onglets :
      - "Creer"  : mode Nouvelle (type/priorite/categorie peuples depuis le serveur)
                   ou Dupliquer (reprend les reglages d'une alarme existante).
                   Portee : toutes les cameras ou une selection ; une alarme globale
                   ou une alarme par camera (modele de nom avec {camera}).
      - "Gerer"  : liste filtrable des alarmes existantes, avec selection multiple pour
                   Activer / Desactiver / Supprimer en masse, plus un bouton de nettoyage
                   des alarmes techniques "__mt_*" laissees par une operation interrompue.
    Les definitions sont creees via New-VmsAlarmDefinition et supprimees via
    Remove-VmsAlarmDefinition (reversibles / gerables depuis le Management Client).
#>

# Extrait le GUID nu d'un item de configuration. Selon le type d'item, le GUID
# se trouve dans .Id OU dans .Path (format "EventTypeGroup[guid]" / "EventType[guid]").
# On teste les deux et on renvoie le premier GUID valide, sinon chaine vide.
function Get-AlarmGuid {
    param([object]$Item)
    $candidates = if ($Item -is [string]) { @($Item) } else { @("$($Item.Id)", "$($Item.Path)") }
    foreach ($candidate in $candidates) {
        if ($candidate -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
            return $Matches[1]
        }
    }
    return ''
}

# Normalise un GUID en minuscules, sans accolades.
function Get-NormGuid {
    param([string]$G)
    try { return ([guid]$G).ToString().ToLowerInvariant() } catch { return $G.ToLowerInvariant() }
}

# Correspondance entre les groupes d'evenement d'ALARME (GUID = constantes systeme
# Milestone, identiques sur toutes les installations) et les groupes de la hierarchie
# de configuration (dont les DisplayName sont en anglais/invariants). Les types
# d'evenement — avec leurs GUID valides pour les alarmes — sont lus dans ces groupes
# folder. Cle = GUID d'alarme (minuscules) ; valeur = prefixes de DisplayName folder.
$script:AlarmGroupFolderMap = @{
    '1eacbcad-d566-4375-834b-cfbe3d937caa' = @('Device')                       # Peripheriques
    '6b90aee7-e6a5-4b5c-82aa-1686c19afe19' = @('Hardware')                     # Materiel
    'b1ca6710-f244-4ce2-8daf-662e342c405a' = @('System events')               # Systeme
    '64d2c24b-a92f-48a0-b9ca-707af7828d67' = @('Recorder')                     # Serveurs d'enregistrement
    '5946b6fa-44d9-4f4c-82bb-46a17b924265' = @('External')                     # Externes
    '2fb8e979-188e-44cc-b1b7-a7ed9c91e5c0' = @('System monitor')              # Moniteur systeme
    'a96692c8-51b1-4f87-b12c-0d3d9cbfc5a4' = @('Analytics')                    # Analytique
    '921f36dc-8d7a-4895-bb9b-265d44f9e784' = @('Access control - Categories')  # Controle d'acces (categories)
    '69d5a520-576c-4040-9fcc-f4442aa21438' = @('Transaction')                  # Transaction
}

# Prefixe des alarmes techniques (probe / preflight) creees puis nettoyees par l'outil.
$script:AlarmTechPrefix = '__mt_'

function Show-AlarmBuilderDialog {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [scriptblock]$Log)

    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Windows.Forms      -ErrorAction SilentlyContinue

    # --- Interrogation du serveur : options reellement disponibles ---
    try {
        $ms = Get-VmsManagementServer

        # Priorites / categories : lisibles sur une definition en attente (SANS ValidateItem)
        $probe = $ms.AlarmDefinitionFolder.AddAlarmDefinition()
        $probe.Name = '__mt_probe__'
        $priorities = @($probe.PriorityValues.Keys | Sort-Object)
        $categories = @($probe.CategoryValues.Keys | Sort-Object)

        # Groupes d'alarme : noms (localises) + GUID attendus par le serveur, lus dans
        # EventTypeGroupValues (fiable, sans ValidateItem).
        $script:_AL_GroupMap = [ordered]@{}      # nom localise -> GUID d'alarme
        foreach ($k in $probe.EventTypeGroupValues.Keys) {
            $script:_AL_GroupMap[$k] = (Get-NormGuid "$($probe.EventTypeGroupValues[$k])")
        }

        # Groupes de la hierarchie de config : source des types d'evenement et de leurs
        # GUID (valides pour les alarmes). On les garde pour resoudre les evenements a la
        # selection du groupe. On EVITE ValidateItem (non supporte sur ce SDK).
        $script:_AL_FolderGroups = @($ms.EventTypeGroupFolder.EventTypeGroups | Where-Object { $_ })
    }
    catch {
        [System.Windows.MessageBox]::Show(
            ($script:T.AL_ErrOptions -f $_.Exception.Message),
            $script:T.AL_Title, 'OK', 'Error') | Out-Null
        return $null
    }

    $script:_AL_Existing = @()
    try { $script:_AL_Existing = @(Get-VmsAlarmDefinition | Sort-Object Name) } catch {}

    $script:_AL_EventMap    = @{}
    $script:_AL_SelCameras  = @()
    $script:_AL_Result      = $null
    $script:_AL_ManageView  = @()

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($script:T.AL_Title)"
        Width="560" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#1E1E2E" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="FontSize"   Value="12"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="FontSize"   Value="12"/>
            <Setter Property="Margin"     Value="0,0,16,0"/>
        </Style>
        <Style TargetType="ListBox">
            <Setter Property="Background"  Value="#313244"/>
            <Setter Property="BorderBrush" Value="#45475A"/>
            <Setter Property="Foreground"  Value="#CDD6F4"/>
            <Setter Property="FontSize"    Value="12"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="Padding"    Value="4,2"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#45475A"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Height"     Value="30"/>
            <Setter Property="FontSize"   Value="12"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Focusable="False" ClickMode="Press"
                                IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border Background="#313244" BorderBrush="#45475A"
                                                BorderThickness="1" CornerRadius="3">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="22"/>
                                                </Grid.ColumnDefinitions>
                                                <ContentPresenter Grid.Column="0" Margin="8,0,0,0"
                                                    Content="{Binding SelectionBoxItem, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                                                    VerticalAlignment="Center"
                                                    TextBlock.Foreground="#CDD6F4" TextBlock.FontSize="12"/>
                                                <Path Grid.Column="1" Data="M 0 0 L 6 6 L 12 0"
                                                      Stroke="#CDD6F4" StrokeThickness="1.5"
                                                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Grid>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <Popup IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom"
                                   AllowsTransparency="True" Focusable="False">
                                <Border Background="#313244" BorderBrush="#45475A" BorderThickness="1"
                                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                                    <ScrollViewer MaxHeight="240" VerticalScrollBarVisibility="Auto">
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="FontSize"   Value="12"/>
            <Setter Property="Padding"    Value="8,5"/>
        </Style>
        <Style x:Key="Lbl" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#A6ADC8"/>
            <Setter Property="FontSize"   Value="11"/>
            <Setter Property="Margin"     Value="0,10,0,3"/>
        </Style>
        <!-- Bouton plat qui respecte le theme sombre, y compris a l'etat desactive
             (le template WPF par defaut affiche sinon un fond gris clair). -->
        <Style x:Key="FlatBtn" TargetType="Button">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="3">
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

    <StackPanel Margin="22,18,22,20">
        <TextBlock Text="$($script:T.AL_Header)" FontSize="16" FontWeight="Bold" Margin="0,0,0,4"/>
        <TextBlock Text="$($script:T.AL_Subtitle)" Foreground="#6C7086" FontSize="11" Margin="0,0,0,10"/>

        <Border Background="#181825" CornerRadius="6" Padding="12,8" Margin="0,0,0,4">
            <StackPanel Orientation="Horizontal">
                <RadioButton x:Name="RbTopCreate" GroupName="TopMode" IsChecked="True"
                             Content="$($script:T.AL_TopCreate)" FontWeight="Bold"/>
                <RadioButton x:Name="RbTopManage" GroupName="TopMode"
                             Content="$($script:T.AL_TopManage)" FontWeight="Bold"/>
            </StackPanel>
        </Border>

        <!-- ============ ONGLET CREER ============ -->
        <StackPanel x:Name="PanelCreate">
            <TextBlock Text="$($script:T.AL_ModeLabel)" Style="{StaticResource Lbl}"/>
            <StackPanel Orientation="Horizontal">
                <RadioButton x:Name="RbNew" GroupName="Mode" IsChecked="True" Content="$($script:T.AL_ModeNew)"/>
                <RadioButton x:Name="RbDup" GroupName="Mode" Content="$($script:T.AL_ModeDup)"/>
            </StackPanel>

            <StackPanel x:Name="PanelNew">
                <TextBlock Text="$($script:T.AL_GroupLabel)" Style="{StaticResource Lbl}"/>
                <ComboBox x:Name="CboGroup"/>
                <TextBlock Text="$($script:T.AL_EventLabel)" Style="{StaticResource Lbl}"/>
                <ComboBox x:Name="CboEvent"/>
                <Grid Margin="0,0,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Text="$($script:T.AL_PriorityLabel)" Style="{StaticResource Lbl}"/>
                        <ComboBox x:Name="CboPriority"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Text="$($script:T.AL_CategoryLabel)" Style="{StaticResource Lbl}"/>
                        <ComboBox x:Name="CboCategory"/>
                    </StackPanel>
                </Grid>
                <TextBlock Text="$($script:T.AL_InstrLabel)" Style="{StaticResource Lbl}"/>
                <TextBox x:Name="TxtInstr" Height="46" TextWrapping="Wrap" AcceptsReturn="True"
                         Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" FontSize="12" Padding="6,4"/>
            </StackPanel>

            <StackPanel x:Name="PanelDup" Visibility="Collapsed">
                <TextBlock Text="$($script:T.AL_DupLabel)" Style="{StaticResource Lbl}"/>
                <ComboBox x:Name="CboExisting"/>
                <TextBlock x:Name="LblDupInfo" Text="" Foreground="#6C7086" FontSize="11" Margin="0,6,0,0" TextWrapping="Wrap"/>
            </StackPanel>

            <Border Background="#181825" CornerRadius="6" Padding="14,10" Margin="0,14,0,0">
                <StackPanel>
                    <TextBlock Text="$($script:T.AL_ScopeLabel)" Foreground="#A6E3A1" FontSize="12" FontWeight="Bold" Margin="0,0,0,6"/>
                    <StackPanel Orientation="Horizontal">
                        <RadioButton x:Name="RbAll" GroupName="Scope" IsChecked="True" Content="$($script:T.AL_ScopeAll)"/>
                        <RadioButton x:Name="RbSel" GroupName="Scope" Content="$($script:T.AL_ScopeSel)"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                        <Button x:Name="BtnPick" Content="$($script:T.AL_PickBtn)" IsEnabled="False"
                                Style="{StaticResource FlatBtn}"
                                Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                                FontSize="11" Padding="10,4" Cursor="Hand"/>
                        <TextBlock x:Name="LblSel" Text="" Foreground="#6C7086" FontSize="11"
                                   VerticalAlignment="Center" Margin="10,0,0,0"/>
                    </StackPanel>
                    <TextBlock Text="$($script:T.AL_GranLabel)" Foreground="#A6E3A1" FontSize="12" FontWeight="Bold" Margin="0,12,0,6"/>
                    <StackPanel Orientation="Horizontal">
                        <RadioButton x:Name="RbGlobal" GroupName="Gran" IsChecked="True" Content="$($script:T.AL_GranGlobal)"/>
                        <RadioButton x:Name="RbPer" GroupName="Gran" Content="$($script:T.AL_GranPer)"/>
                    </StackPanel>
                </StackPanel>
            </Border>

            <TextBlock x:Name="LblName" Text="$($script:T.AL_NameLabel)" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtName" Height="30"
                     Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" FontSize="12"
                     Padding="8,0" VerticalContentAlignment="Center"/>
            <TextBlock x:Name="LblNameHint" Text="" Foreground="#6C7086" FontSize="10" Margin="0,3,0,0"/>
        </StackPanel>

        <!-- ============ ONGLET GERER ============ -->
        <StackPanel x:Name="PanelManage" Visibility="Collapsed">
            <TextBlock Text="$($script:T.AL_MngHeader)" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="TxtFilter" Height="28"
                     Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" FontSize="12"
                     Padding="8,0" VerticalContentAlignment="Center"/>
            <ListBox x:Name="LstAlarms" Height="220" Margin="0,8,0,0"
                     SelectionMode="Extended" BorderThickness="1"/>
            <Grid Margin="0,8,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="LblMngCount" Grid.Column="0" Text="" Foreground="#6C7086"
                           FontSize="11" VerticalAlignment="Center"/>
                <Button x:Name="BtnMngSelectAll" Grid.Column="1" Content="$($script:T.AL_MngSelectAll)"
                        Style="{StaticResource FlatBtn}"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="11" Padding="10,4" Cursor="Hand"/>
            </Grid>
            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                <Button x:Name="BtnMngEnable" Content="$($script:T.AL_MngEnable)"
                        Style="{StaticResource FlatBtn}"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="12" Padding="14,6" Cursor="Hand" Margin="0,0,8,0"/>
                <Button x:Name="BtnMngDisable" Content="$($script:T.AL_MngDisable)"
                        Style="{StaticResource FlatBtn}"
                        Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                        FontSize="12" Padding="14,6" Cursor="Hand" Margin="0,0,8,0"/>
                <Button x:Name="BtnMngDelete" Content="$($script:T.AL_MngDelete)"
                        Background="#F38BA8" Foreground="#1E1E2E" BorderThickness="0"
                        FontWeight="Bold" FontSize="12" Padding="14,6" Cursor="Hand"/>
            </StackPanel>
            <Border BorderBrush="#313244" BorderThickness="0,1,0,0" Margin="0,14,0,0" Padding="0,10,0,0">
                <StackPanel>
                    <Button x:Name="BtnMngClean" Content="$($script:T.AL_MngClean)" HorizontalAlignment="Left"
                            Style="{StaticResource FlatBtn}"
                            Background="#313244" Foreground="#F9E2AF" BorderBrush="#45475A" BorderThickness="1"
                            FontSize="11" Padding="10,4" Cursor="Hand"/>
                    <TextBlock Text="$($script:T.AL_MngCleanHint)" Foreground="#6C7086" FontSize="10"
                               Margin="0,4,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>
        </StackPanel>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="BtnCancel" Content="$($script:T.AL_BtnCancel)" IsCancel="True"
                    Background="#313244" Foreground="#CDD6F4" BorderBrush="#45475A" BorderThickness="1"
                    FontSize="13" Padding="20,8" Cursor="Hand" Margin="0,0,10,0"/>
            <Button x:Name="BtnCreate" Content="$($script:T.AL_BtnCreate)" IsDefault="True"
                    Background="#A6E3A1" Foreground="#1E1E2E" BorderThickness="0"
                    FontWeight="Bold" FontSize="13" Padding="20,8" Cursor="Hand"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $win    = [System.Windows.Markup.XamlReader]::Load($reader)

    $ctl = {param($n) $win.FindName($n)}
    $RbTopCreate=&$ctl RbTopCreate; $RbTopManage=&$ctl RbTopManage
    $PanelCreate=&$ctl PanelCreate; $PanelManage=&$ctl PanelManage
    $RbNew=&$ctl RbNew; $RbDup=&$ctl RbDup; $PanelNew=&$ctl PanelNew; $PanelDup=&$ctl PanelDup
    $CboGroup=&$ctl CboGroup; $CboEvent=&$ctl CboEvent; $CboPriority=&$ctl CboPriority
    $CboCategory=&$ctl CboCategory; $TxtInstr=&$ctl TxtInstr; $CboExisting=&$ctl CboExisting
    $LblDupInfo=&$ctl LblDupInfo; $RbAll=&$ctl RbAll; $RbSel=&$ctl RbSel; $BtnPick=&$ctl BtnPick
    $LblSel=&$ctl LblSel; $RbGlobal=&$ctl RbGlobal; $RbPer=&$ctl RbPer
    $LblName=&$ctl LblName; $TxtName=&$ctl TxtName; $LblNameHint=&$ctl LblNameHint
    $BtnCreate=&$ctl BtnCreate; $BtnCancel=&$ctl BtnCancel
    $TxtFilter=&$ctl TxtFilter; $LstAlarms=&$ctl LstAlarms; $LblMngCount=&$ctl LblMngCount
    $BtnMngSelectAll=&$ctl BtnMngSelectAll; $BtnMngEnable=&$ctl BtnMngEnable
    $BtnMngDisable=&$ctl BtnMngDisable; $BtnMngDelete=&$ctl BtnMngDelete; $BtnMngClean=&$ctl BtnMngClean

    # Remplissage des listes (creation)
    foreach ($g in $script:_AL_GroupMap.Keys) { [void]$CboGroup.Items.Add($g) }
    [void]$CboPriority.Items.Add($script:T.AL_DefaultOpt)
    foreach ($p in $priorities) { [void]$CboPriority.Items.Add($p) }
    $CboPriority.SelectedIndex = 0
    [void]$CboCategory.Items.Add($script:T.AL_NoneOpt)
    foreach ($c in $categories) { [void]$CboCategory.Items.Add($c) }
    $CboCategory.SelectedIndex = 0
    foreach ($a in $script:_AL_Existing) { [void]$CboExisting.Items.Add($a.Name) }

    # Cascade groupe -> types d'evenement.
    # On mappe le GUID (constant) du groupe d'alarme vers les groupes folder anglais
    # correspondants, et on agrege leurs types d'evenement (nom + GUID via .Path).
    $CboGroup.Add_SelectionChanged({
        $CboEvent.Items.Clear()
        $script:_AL_EventMap = [ordered]@{}
        $g = $CboGroup.SelectedItem
        if (-not $g) { return }
        try {
            $alarmGuid = $script:_AL_GroupMap[$g]
            $patterns  = $script:AlarmGroupFolderMap[$alarmGuid]
            if ($patterns) {
                foreach ($fg in $script:_AL_FolderGroups) {
                    $dn = "$($fg.DisplayName)"
                    $isMatch = $false
                    foreach ($p in $patterns) { if ($dn -like "$p*") { $isMatch = $true; break } }
                    if (-not $isMatch) { continue }
                    foreach ($et in $fg.EventTypeFolder.EventTypes) {
                        if ($null -eq $et) { continue }
                        $guid = Get-AlarmGuid $et
                        if (-not $guid) { continue }
                        $en = if ($et.DisplayName) { $et.DisplayName } else { $et.Name }
                        if (-not $script:_AL_EventMap.Contains($en)) { $script:_AL_EventMap[$en] = $guid }
                    }
                }
            }
            foreach ($en in ($script:_AL_EventMap.Keys | Sort-Object)) { [void]$CboEvent.Items.Add($en) }
            if ($CboEvent.Items.Count -gt 0) {
                $CboEvent.SelectedIndex = 0
            } else {
                # Groupe sans correspondance connue : on informe plutot que d'echouer en silence
                $LblNameHint.Text = $script:T.AL_NoEventsForGroup
            }
        } catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.AL_ErrEvents -f $_.Exception.Message),
                $script:T.AL_Title, 'OK', 'Warning') | Out-Null
        }
    })

    # Bascule Nouvelle / Dupliquer
    $updateMode = {
        if ($RbDup.IsChecked) {
            $PanelNew.Visibility = 'Collapsed'; $PanelDup.Visibility = 'Visible'
        } else {
            $PanelNew.Visibility = 'Visible'; $PanelDup.Visibility = 'Collapsed'
        }
    }
    $RbNew.Add_Checked($updateMode)
    $RbDup.Add_Checked($updateMode)

    $CboExisting.Add_SelectionChanged({
        $i = $CboExisting.SelectedIndex
        if ($i -lt 0) { return }
        $a = $script:_AL_Existing[$i]
        $LblDupInfo.Text = $script:T.AL_DupInfo -f $a.Name
    })

    # Portee : activer le choix des cameras
    $updateScope = {
        $BtnPick.IsEnabled = [bool]$RbSel.IsChecked
        if ($RbAll.IsChecked) { $LblSel.Text = '' }
    }
    $RbAll.Add_Checked($updateScope)
    $RbSel.Add_Checked($updateScope)

    $BtnPick.Add_Click({
        $cams = @(Select-Camera -Title $script:T.AL_PickTitle -RemoveDuplicates)
        if ($cams.Count -gt 0) {
            $script:_AL_SelCameras = $cams
            $LblSel.Text = $script:T.AL_SelCount -f $cams.Count
        }
    })

    # Indice du champ nom selon la granularite
    $updateName = {
        if ($RbPer.IsChecked) {
            $LblName.Text     = $script:T.AL_NameTemplateLabel
            $LblNameHint.Text = $script:T.AL_NameHint
        } else {
            $LblName.Text     = $script:T.AL_NameLabel
            $LblNameHint.Text = ''
        }
    }
    $RbGlobal.Add_Checked($updateName)
    $RbPer.Add_Checked($updateName)

    # ---------- Logique de l'onglet GERER ----------
    $reloadAlarms = {
        try { $script:_AL_Existing = @(Get-VmsAlarmDefinition | Sort-Object Name) }
        catch { $script:_AL_Existing = @() }
    }
    $updateMngCount = {
        $LblMngCount.Text = $script:T.AL_MngCount -f $script:_AL_ManageView.Count, $LstAlarms.SelectedItems.Count
    }
    # Repeuple la liste selon le filtre. Chaque item est un ListBoxItem dont .Tag porte
    # l'objet AlarmDefinition (mapping fiable meme si des noms sont identiques).
    $refreshManage = {
        $filter = "$($TxtFilter.Text)".Trim().ToLowerInvariant()
        $LstAlarms.Items.Clear()
        $view = [System.Collections.Generic.List[object]]::new()
        foreach ($a in $script:_AL_Existing) {
            $nm = "$($a.Name)"
            if ($filter -and ($nm.ToLowerInvariant().IndexOf($filter) -lt 0)) { continue }
            $view.Add($a)
            $enabled = $true
            try { $enabled = [bool]$a.Enabled } catch {}
            $dot = if ($enabled) { [char]0x25CF } else { [char]0x25CB }
            $tag = if ($enabled) { '' } else { $script:T.AL_MngDisabledTag }
            $item = New-Object System.Windows.Controls.ListBoxItem
            $item.Content = "$dot  $nm$tag"
            $item.Tag     = $a
            [void]$LstAlarms.Items.Add($item)
        }
        $script:_AL_ManageView = $view
        if ($view.Count -eq 0) { $LblMngCount.Text = $script:T.AL_MngEmpty }
        else { & $updateMngCount }
    }
    # Applique une operation (scriptblock) a chaque alarme selectionnee, compte OK/erreurs,
    # journalise via le modele $LogKey, puis recharge la liste depuis le serveur.
    $applyToSelection = {
        param([scriptblock]$Op, [string]$LogKey)
        $sel = [System.Collections.Generic.List[object]]::new()
        foreach ($it in $LstAlarms.SelectedItems) { $sel.Add($it.Tag) }
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show($script:T.AL_MngNoSel, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
            return $null
        }
        $ok = 0; $err = 0
        foreach ($a in $sel) {
            try { & $Op $a; $ok++ } catch { $err++ }
        }
        & $Log ($script:T[$LogKey] -f $ok, $err)
        & $reloadAlarms
        & $refreshManage
    }

    $TxtFilter.Add_TextChanged({ & $refreshManage })
    $LstAlarms.Add_SelectionChanged({
        if ($script:_AL_ManageView.Count -gt 0) { & $updateMngCount }
    })
    $BtnMngSelectAll.Add_Click({ $LstAlarms.SelectAll() })

    $BtnMngEnable.Add_Click({
        & $applyToSelection { param($a) $a.Enabled = $true;  $a.Save() } 'AL_MngLogEnable'
    })
    $BtnMngDisable.Add_Click({
        & $applyToSelection { param($a) $a.Enabled = $false; $a.Save() } 'AL_MngLogDisable'
    })
    $BtnMngDelete.Add_Click({
        $count = $LstAlarms.SelectedItems.Count
        if ($count -eq 0) {
            [System.Windows.MessageBox]::Show($script:T.AL_MngNoSel, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
            return
        }
        $r = [System.Windows.MessageBox]::Show(
            ($script:T.AL_MngConfirmDel -f $count),
            $script:T.AL_MngConfirmTitle, 'YesNo', 'Warning')
        if ($r -ne 'Yes') { return }
        & $applyToSelection { param($a) Remove-VmsAlarmDefinition -AlarmDefinition $a -Confirm:$false -ErrorAction Stop } 'AL_MngLogDel'
    })
    $BtnMngClean.Add_Click({
        & $reloadAlarms
        $junk = @($script:_AL_Existing | Where-Object { "$($_.Name)" -like "$($script:AlarmTechPrefix)*" })
        $n = 0
        foreach ($a in $junk) {
            try { Remove-VmsAlarmDefinition -AlarmDefinition $a -Confirm:$false -ErrorAction Stop; $n++ } catch {}
        }
        & $Log ($script:T.AL_MngLogClean -f $n)
        & $reloadAlarms
        & $refreshManage
    })

    # Bascule Creer / Gerer
    $updateTopMode = {
        if ($RbTopManage.IsChecked) {
            $PanelCreate.Visibility = 'Collapsed'; $PanelManage.Visibility = 'Visible'
            $BtnCreate.Visibility   = 'Collapsed'
            $BtnCancel.Content      = $script:T.AL_BtnClose
            & $reloadAlarms
            & $refreshManage
        } else {
            $PanelCreate.Visibility = 'Visible'; $PanelManage.Visibility = 'Collapsed'
            $BtnCreate.Visibility   = 'Visible'
            $BtnCancel.Content      = $script:T.AL_BtnCancel
        }
    }
    $RbTopCreate.Add_Checked($updateTopMode)
    $RbTopManage.Add_Checked($updateTopMode)

    $BtnCreate.Add_Click({
        # Validation portee
        if ($RbSel.IsChecked -and $script:_AL_SelCameras.Count -eq 0) {
            [System.Windows.MessageBox]::Show($script:T.AL_ValNoCams, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($TxtName.Text)) {
            [System.Windows.MessageBox]::Show($script:T.AL_ValNoName, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
            return
        }

        $groupId = $null; $eventId = $null; $priority = $null; $category = $null; $instr = ''

        if ($RbDup.IsChecked) {
            $i = $CboExisting.SelectedIndex
            if ($i -lt 0) {
                [System.Windows.MessageBox]::Show($script:T.AL_ValNoDup, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
                return
            }
            $src = $script:_AL_Existing[$i]
            $groupId = Get-AlarmGuid $src.EventTypeGroup
            $eventId = Get-AlarmGuid $src.EventType
            $priority = ($src.PriorityValues.GetEnumerator() | Where-Object { $_.Value -eq $src.Priority } | Select-Object -First 1).Key
            $category = ($src.CategoryValues.GetEnumerator() | Where-Object { $_.Value -eq $src.Category } | Select-Object -First 1).Key
            $instr = $src.Description
        }
        else {
            $g = $CboGroup.SelectedItem
            $e = $CboEvent.SelectedItem
            if (-not $g -or -not $e) {
                [System.Windows.MessageBox]::Show($script:T.AL_ValNoEvent, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
                return
            }
            $groupId = $script:_AL_GroupMap[$g]
            $eventId = $script:_AL_EventMap[$e]
            if ($CboPriority.SelectedIndex -gt 0) { $priority = $CboPriority.SelectedItem }
            if ($CboCategory.SelectedIndex -gt 0) { $category = $CboCategory.SelectedItem }
            $instr = $TxtInstr.Text
        }

        if ([string]::IsNullOrWhiteSpace($groupId) -or [string]::IsNullOrWhiteSpace($eventId)) {
            [System.Windows.MessageBox]::Show($script:T.AL_ErrNoGuid, $script:T.AL_Title, 'OK', 'Warning') | Out-Null
            return
        }

        $script:_AL_Result = @{
            GroupId      = "$groupId"
            EventId      = "$eventId"
            Priority     = $priority
            Category     = $category
            Instructions = $instr
            Scope        = if ($RbSel.IsChecked) { 'Selection' } else { 'AllCameras' }
            Cameras      = $script:_AL_SelCameras
            PerCamera    = [bool]$RbPer.IsChecked
            Name         = $TxtName.Text
        }
        $win.DialogResult = $true
    })

    if ($CboGroup.Items.Count -gt 0) { $CboGroup.SelectedIndex = 0 }
    & $updateName

    if ($win.ShowDialog() -eq $true) { return $script:_AL_Result }
    return $null
}


function New-BulkAlarm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [hashtable]$Config,
        [Parameter(Mandatory)] [scriptblock]$Log,
        [Parameter()] [scriptblock]$Cancel = { $false },
        [Parameter()] [scriptblock]$ReportProgress = {}
    )

    $cfg = Show-AlarmBuilderDialog -Log $Log
    if (-not $cfg) { & $Log $script:T.AL_Cancelled ; return }

    function New-OneAlarm {
        param([string]$AlarmName, [string[]]$SourcePaths, $RelatedCams)
        $p = @{
            Name           = $AlarmName
            EventTypeGroup = $cfg.GroupId
            EventType      = $cfg.EventId
            Source         = $SourcePaths
            ErrorAction    = 'Stop'
        }
        if ($cfg.Priority)                              { $p.Priority       = $cfg.Priority }
        if ($cfg.Category)                              { $p.Category       = $cfg.Category }
        if (-not [string]::IsNullOrWhiteSpace($cfg.Instructions)) { $p.Instructions = $cfg.Instructions }
        if ($RelatedCams)                               { $p.RelatedCameras = $RelatedCams }
        New-VmsAlarmDefinition @p | Out-Null
    }

    # --- Pre-vol : certains types d'evenement ne sont pas des declencheurs d'alarme
    # valides (le serveur repond "Triggering event type is out of range"). On ne peut
    # pas connaitre la liste valide a l'avance (EventTypeValues indisponible sur ce SDK),
    # donc on cree une alarme jetable pour valider le choix AVANT la creation en masse.
    $preflightName = "$($script:AlarmTechPrefix)preflight_$([guid]::NewGuid().ToString('N').Substring(0,8))"
    try {
        New-OneAlarm -AlarmName $preflightName -SourcePaths @('AllCameras') -RelatedCams $null
    }
    catch {
        & $Log ($script:T.AL_PreflightFail -f $_.Exception.Message)
        return
    }
    finally {
        # Nettoyage de l'alarme de test (qu'elle ait reussi ou non)
        try {
            Get-VmsAlarmDefinition | Where-Object { $_.Name -eq $preflightName } | ForEach-Object {
                try { Remove-VmsAlarmDefinition -AlarmDefinition $_ -Confirm:$false -ErrorAction Stop }
                catch { $_ | Remove-VmsAlarmDefinition -Confirm:$false -ErrorAction SilentlyContinue }
            }
        } catch {}
    }

    if ($cfg.PerCamera) {
        $cams = if ($cfg.Scope -eq 'AllCameras') { @(Get-VmsCamera) } else { @($cfg.Cameras) }
        $total = $cams.Count
        & $Log ($script:T.AL_LogPerCam -f $total)

        $done = 0; $ok = 0; $err = 0
        foreach ($cam in $cams) {
            if (& $Cancel) { & $Log ($script:T.AL_LogCancelled -f $done, $total) ; break }
            $done++
            & $ReportProgress $done $total
            $name = if ($cfg.Name -match '\{camera\}') {
                $cfg.Name -replace '\{camera\}', $cam.Name
            } else {
                "$($cfg.Name) - $($cam.Name)"
            }
            try   { New-OneAlarm -AlarmName $name -SourcePaths @($cam.Path) -RelatedCams @($cam) ; $ok++ ; & $Log ($script:T.AL_LogOk -f $name) }
            catch { $err++ ; & $Log ($script:T.AL_LogErr -f $name, $_.Exception.Message) }
        }
        & $Log ($script:T.AL_LogDonePer -f $ok, $err)
    }
    else {
        if ($cfg.Scope -eq 'AllCameras') {
            $sourcePaths = @('AllCameras') ; $rel = $null
        } else {
            $sourcePaths = @($cfg.Cameras | ForEach-Object { $_.Path }) ; $rel = @($cfg.Cameras)
        }
        try {
            New-OneAlarm -AlarmName $cfg.Name -SourcePaths $sourcePaths -RelatedCams $rel
            & $Log ($script:T.AL_LogOk -f $cfg.Name)
            & $Log $script:T.AL_LogDoneGlobal
        }
        catch { & $Log ($script:T.AL_LogErr -f $cfg.Name, $_.Exception.Message) }
    }
}
