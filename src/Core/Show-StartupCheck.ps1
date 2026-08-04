<#
.SYNOPSIS
    Fenetre GUI de verification des dependances au demarrage.
#>

function Show-StartupCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AppRoot
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $script:_SC_Result      = $false
    $script:_SC_AppRoot     = $AppRoot
    $script:_SC_DepsPath    = Join-Path $AppRoot 'Dependencies'
    $script:_SC_IsOffline   = Test-Path $script:_SC_DepsPath
    $script:_SC_DepRows     = @{}
    $script:_SC_Modules     = Get-RequiredModules | ForEach-Object {
        @{ Name = $_.Name; Required = $_.Required; Description = $script:T[$_.DescriptionKey] }
    }

    # SECURITE : depot de mise a jour CODE EN DUR. Toute valeur differente dans
    # config.json est rejetee (un fichier local modifie ne doit jamais rediriger
    # les mises a jour vers un depot tiers => execution de code arbitraire).
    $script:_SC_TrustedRepo = 'Nothinx-44/XProtect-Export-Tool-to-Excel-MilestonePSTools-GUI-'

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Milestone Toolkit"
        Width="580" Height="620"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E2E"
        FontFamily="Segoe UI">

    <Window.Resources>
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="#89B4FA"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#B4D0FF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#7AA2F7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="Transparent"/>
            <Setter Property="Foreground"      Value="#A6ADC8"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush"     Value="#45475A"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#313244"/>
                                <Setter Property="Foreground" Value="#CDD6F4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="WarnBtn" TargetType="Button">
            <Setter Property="Background"      Value="#F9A825"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="20,10"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#FFCA28"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#F57F17"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="GreenBtn" TargetType="Button">
            <Setter Property="Background"      Value="#A6E3A1"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="14,8"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#C3F0BE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#89D4A1"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BadgeBtn" TargetType="Button">
            <Setter Property="Background"      Value="#89B4FA"/>
            <Setter Property="Foreground"      Value="#1E1E2E"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="12,6"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#B4D0FF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#7AA2F7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#181825" Padding="32,24,32,20">
            <StackPanel Orientation="Horizontal">
                <Border Width="38" Height="38" Background="#89B4FA" CornerRadius="8"
                        Margin="0,0,14,0" VerticalAlignment="Center">
                    <TextBlock Text="M" FontSize="22" FontWeight="Bold"
                               Foreground="#1E1E2E"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Milestone Toolkit"
                               FontSize="20" FontWeight="Bold" Foreground="#CDD6F4"/>
                    <TextBlock x:Name="HeaderSubtitle"
                               Text="Verification des dependances au demarrage"
                               FontSize="12" Foreground="#6C7086" Margin="0,3,0,0"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="24,20,24,0">
            <StackPanel x:Name="DepsPanel"/>
        </ScrollViewer>

        <Border x:Name="OfflineBanner" Grid.Row="2"
                Background="#2A2A1A" BorderBrush="#F9A825" BorderThickness="0,0,0,0"
                CornerRadius="6" Margin="24,12,24,0" Padding="16,12"
                Visibility="Collapsed">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock x:Name="OfflineBannerTitle"
                               Text="" FontSize="12" FontWeight="SemiBold" Foreground="#F9A825"/>
                    <TextBlock x:Name="OfflineBannerText"
                               Text="" FontSize="11" Foreground="#A6ADC8"
                               TextWrapping="Wrap" Margin="0,3,0,0"/>
                </StackPanel>
                <Button x:Name="BtnSaveDeps" Grid.Column="1"
                        Content="" Margin="12,0,0,0"
                        Style="{StaticResource GreenBtn}" Visibility="Collapsed"/>
            </Grid>
        </Border>

        <StackPanel Grid.Row="3" Margin="24,12,24,8">
            <ProgressBar x:Name="ProgressBar"
                         Height="3" IsIndeterminate="True"
                         Background="#313244" Foreground="#89B4FA"
                         BorderThickness="0" Visibility="Collapsed" Margin="0,0,0,10"/>
            <TextBlock x:Name="StatusText"
                       Text="" FontSize="12" Foreground="#A6ADC8" TextWrapping="Wrap"/>
        </StackPanel>

        <Border Grid.Row="4" Background="#181825" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Made by Vincent Le Bonhomme"
                           FontSize="10" Foreground="#45475A"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnQuit"    Content=""
                            Style="{StaticResource SecondaryBtn}" Margin="0,0,10,0"/>
                    <Button x:Name="BtnInstall" Content=""
                            Style="{StaticResource WarnBtn}"
                            Visibility="Collapsed" Margin="0,0,10,0"/>
                    <Button x:Name="BtnLaunch"  Content=""
                            Style="{StaticResource PrimaryBtn}" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $script:_SC_Win         = [System.Windows.Markup.XamlReader]::Load($reader)

    $script:_SC_DepsPanel          = $script:_SC_Win.FindName('DepsPanel')
    $script:_SC_Progress           = $script:_SC_Win.FindName('ProgressBar')
    $script:_SC_Status             = $script:_SC_Win.FindName('StatusText')
    $script:_SC_BtnInstall         = $script:_SC_Win.FindName('BtnInstall')
    $script:_SC_BtnLaunch          = $script:_SC_Win.FindName('BtnLaunch')
    $script:_SC_BtnQuit            = $script:_SC_Win.FindName('BtnQuit')
    $script:_SC_BtnSaveDeps        = $script:_SC_Win.FindName('BtnSaveDeps')
    $script:_SC_OfflineBanner      = $script:_SC_Win.FindName('OfflineBanner')
    $script:_SC_OfflineBannerTitle = $script:_SC_Win.FindName('OfflineBannerTitle')
    $script:_SC_OfflineBannerText  = $script:_SC_Win.FindName('OfflineBannerText')

    # Textes traduits
    $script:_SC_Win.Title                           = $script:T.SC_WindowTitle
    $script:_SC_Win.FindName('HeaderSubtitle').Text = $script:T.SC_Subtitle
    $script:_SC_BtnQuit.Content                     = $script:T.SC_BtnQuit
    $script:_SC_BtnInstall.Content                  = $script:T.SC_BtnInstall
    $script:_SC_BtnLaunch.Content                   = $script:T.SC_BtnLaunch
    $script:_SC_BtnSaveDeps.Content                 = $script:T.SC_BtnSaveDeps
    $script:_SC_Status.Text                         = $script:T.SC_StatusInit

    foreach ($mod in $script:_SC_Modules) {
        $name = $mod.Name
        $desc = $mod.Description

        $indicator = [System.Windows.Shapes.Ellipse]::new()
        $indicator.Width  = 12
        $indicator.Height = 12
        $indicator.Fill   = [System.Windows.Media.Brushes]::Gray
        $indicator.VerticalAlignment = 'Center'
        $indicator.Margin = [System.Windows.Thickness]::new(0,0,14,0)

        $lblName = [System.Windows.Controls.TextBlock]::new()
        $lblName.Text       = $name
        $lblName.FontSize   = 13
        $lblName.FontWeight = [System.Windows.FontWeights]::SemiBold
        $lblName.Foreground = [System.Windows.Media.Brushes]::White

        $lblDesc = [System.Windows.Controls.TextBlock]::new()
        $lblDesc.Text       = $desc
        $lblDesc.FontSize   = 11
        $lblDesc.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(108,112,134))
        $lblDesc.Margin     = [System.Windows.Thickness]::new(0,3,0,0)

        $nameStack = [System.Windows.Controls.StackPanel]::new()
        $nameStack.VerticalAlignment = 'Center'
        [void]$nameStack.Children.Add($lblName)
        [void]$nameStack.Children.Add($lblDesc)

        $lblStatus = [System.Windows.Controls.TextBlock]::new()
        $lblStatus.Text              = $script:T.SC_StatusWaiting
        $lblStatus.FontSize          = 12
        $lblStatus.Foreground        = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))
        $lblStatus.VerticalAlignment = 'Center'

        $rowGrid = [System.Windows.Controls.Grid]::new()
        $c0 = [System.Windows.Controls.ColumnDefinition]::new(); $c0.Width = [System.Windows.GridLength]::Auto
        $c1 = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
        $c2 = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::Auto
        [void]$rowGrid.ColumnDefinitions.Add($c0)
        [void]$rowGrid.ColumnDefinitions.Add($c1)
        [void]$rowGrid.ColumnDefinitions.Add($c2)
        [System.Windows.Controls.Grid]::SetColumn($indicator, 0)
        [System.Windows.Controls.Grid]::SetColumn($nameStack,  1)
        [System.Windows.Controls.Grid]::SetColumn($lblStatus,  2)
        [void]$rowGrid.Children.Add($indicator)
        [void]$rowGrid.Children.Add($nameStack)
        [void]$rowGrid.Children.Add($lblStatus)

        $card = [System.Windows.Controls.Border]::new()
        $card.Background   = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(24,24,37))
        $card.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $card.Padding      = [System.Windows.Thickness]::new(16,14,16,14)
        $card.Margin       = [System.Windows.Thickness]::new(0,0,0,10)
        $card.Child        = $rowGrid
        [void]$script:_SC_DepsPanel.Children.Add($card)

        $script:_SC_DepRows[$name] = @{ Indicator = $indicator; Status = $lblStatus; Available = $false }
    }

    # Separateur visuel avant la carte application
    $appSep            = [System.Windows.Controls.Border]::new()
    $appSep.Height     = 1
    $appSep.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(49,50,68))
    $appSep.Margin     = [System.Windows.Thickness]::new(0,4,0,10)
    [void]$script:_SC_DepsPanel.Children.Add($appSep)

    # Indicateur (cercle colore)
    $script:_SC_AppIndicator                   = [System.Windows.Shapes.Ellipse]::new()
    $script:_SC_AppIndicator.Width             = 12
    $script:_SC_AppIndicator.Height            = 12
    $script:_SC_AppIndicator.Fill              = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(108,112,134))
    $script:_SC_AppIndicator.VerticalAlignment = 'Center'
    $script:_SC_AppIndicator.Margin            = [System.Windows.Thickness]::new(0,0,14,0)

    $appLblName            = [System.Windows.Controls.TextBlock]::new()
    $appLblName.Text       = 'Milestone Toolkit'
    $appLblName.FontSize   = 13
    $appLblName.FontWeight = [System.Windows.FontWeights]::SemiBold
    $appLblName.Foreground = [System.Windows.Media.Brushes]::White

    $appLblDesc            = [System.Windows.Controls.TextBlock]::new()
    $appLblDesc.Text       = "$($script:T.SC_AppLabel) $([char]0x2014) v$($script:AppVersion)"
    $appLblDesc.FontSize   = 11
    $appLblDesc.Foreground = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(108,112,134))
    $appLblDesc.Margin     = [System.Windows.Thickness]::new(0,3,0,0)

    $appNameStack                   = [System.Windows.Controls.StackPanel]::new()
    $appNameStack.VerticalAlignment = 'Center'
    [void]$appNameStack.Children.Add($appLblName)
    [void]$appNameStack.Children.Add($appLblDesc)

    $script:_SC_AppStatus                      = [System.Windows.Controls.TextBlock]::new()
    $script:_SC_AppStatus.Text                 = $script:T.SC_AppVerChecking
    $script:_SC_AppStatus.FontSize             = 12
    $script:_SC_AppStatus.Foreground           = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(166,173,200))
    $script:_SC_AppStatus.VerticalAlignment    = 'Center'

    $script:_SC_AppUpdateBtn            = [System.Windows.Controls.Button]::new()
    $script:_SC_AppUpdateBtn.Visibility = 'Collapsed'

    $appRightPanel                   = [System.Windows.Controls.StackPanel]::new()
    $appRightPanel.Orientation       = 'Horizontal'
    $appRightPanel.VerticalAlignment = 'Center'
    [void]$appRightPanel.Children.Add($script:_SC_AppStatus)
    [void]$appRightPanel.Children.Add($script:_SC_AppUpdateBtn)

    $appRowGrid = [System.Windows.Controls.Grid]::new()
    $ac0 = [System.Windows.Controls.ColumnDefinition]::new(); $ac0.Width = [System.Windows.GridLength]::Auto
    $ac1 = [System.Windows.Controls.ColumnDefinition]::new(); $ac1.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
    $ac2 = [System.Windows.Controls.ColumnDefinition]::new(); $ac2.Width = [System.Windows.GridLength]::Auto
    [void]$appRowGrid.ColumnDefinitions.Add($ac0)
    [void]$appRowGrid.ColumnDefinitions.Add($ac1)
    [void]$appRowGrid.ColumnDefinitions.Add($ac2)
    [System.Windows.Controls.Grid]::SetColumn($script:_SC_AppIndicator, 0)
    [System.Windows.Controls.Grid]::SetColumn($appNameStack,             1)
    [System.Windows.Controls.Grid]::SetColumn($appRightPanel,            2)
    [void]$appRowGrid.Children.Add($script:_SC_AppIndicator)
    [void]$appRowGrid.Children.Add($appNameStack)
    [void]$appRowGrid.Children.Add($appRightPanel)

    $appCard              = [System.Windows.Controls.Border]::new()
    $appCard.Background   = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(24,24,37))
    $appCard.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $appCard.Padding      = [System.Windows.Thickness]::new(16,14,16,14)
    $appCard.Child        = $appRowGrid
    [void]$script:_SC_DepsPanel.Children.Add($appCard)

    $script:_SC_AppRelease = $null

    $script:_SC_Refresh = {
        $script:_SC_Win.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Render, [Action]{}
        )
    }

    # Style applique apres creation de la fenetre (ressource XAML disponible a ce stade)
    $script:_SC_AppUpdateBtn.Style = $script:_SC_Win.FindResource('BadgeBtn')

    $script:_SC_CheckAppUpdate = {
        $cfgPath = Join-Path $script:_SC_AppRoot 'config.json'
        $cfg = $null
        if (Test-Path $cfgPath) {
            try { $cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
        }
        # Mise a jour desactivee explicitement dans config.json
        if ($cfg -and $cfg.autoUpdate -and $cfg.autoUpdate.enabled -eq $false) {
            $script:_SC_AppStatus.Text = $script:T.SC_AppVerUpToDate
            & $script:_SC_Refresh
            return
        }

        # SECURITE : on ignore la valeur repo de config.json si elle differe du depot
        # de confiance. Un config.json altere ne doit pas rediriger les mises a jour.
        $configuredRepo = if ($cfg -and $cfg.autoUpdate) { $cfg.autoUpdate.repo } else { $null }
        if ($configuredRepo -and $configuredRepo -ne $script:_SC_TrustedRepo) {
            $script:_SC_AppStatus.Text    = $script:T.SC_AppVerRepoMismatch
            $script:_SC_AppIndicator.Fill = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
            & $script:_SC_Refresh
            return
        }

        try {
            [Net.ServicePointManager]::SecurityProtocol =
                [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $headers = @{ 'User-Agent' = 'MilestoneToolkitUpdater' }
            $release = Invoke-RestMethod `
                -Uri "https://api.github.com/repos/$($script:_SC_TrustedRepo)/releases/latest" `
                -Headers $headers -TimeoutSec 10 -ErrorAction Stop
        }
        catch {
            $script:_SC_AppStatus.Text = $script:T.SC_AppVerNetErr
            $script:_SC_AppIndicator.Fill = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(108,112,134))
            & $script:_SC_Refresh
            return
        }

        $remoteClean  = $release.tag_name -replace '^v',''
        $currentClean = "$($script:AppVersion)"  -replace '^v',''
        $remote  = try { [version]$remoteClean  } catch { $null }
        $current = try { [version]$currentClean } catch { $null }

        if (-not $remote -or -not $current) {
            $script:_SC_AppStatus.Text = $script:T.SC_AppVerNetErr
            & $script:_SC_Refresh
            return
        }

        if ($remote -le $current) {
            $script:_SC_AppStatus.Text      = $script:T.SC_AppVerUpToDate
            $script:_SC_AppStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(166,227,161))
            $script:_SC_AppIndicator.Fill   = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(166,227,161))
        }
        else {
            $script:_SC_AppRelease              = $release
            $script:_SC_AppStatus.Visibility    = 'Collapsed'
            $script:_SC_AppUpdateBtn.Content    = $script:T.SC_AppVerAvailable -f $release.tag_name
            $script:_SC_AppUpdateBtn.Visibility = 'Visible'
            $script:_SC_AppIndicator.Fill       = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(137,180,250))
        }
        & $script:_SC_Refresh
    }

    $script:_SC_SetStatus = {
        param([string]$Name, [string]$State, [string]$Label)
        $row = $script:_SC_DepRows[$Name]
        $row.Status.Text = $Label
        $color = switch ($State) {
            'checking'   { [System.Windows.Media.Color]::FromRgb(249,168, 37) }
            'ok'         { [System.Windows.Media.Color]::FromRgb(166,227,161) }
            'missing'    { [System.Windows.Media.Color]::FromRgb(243,139,168) }
            'installing' { [System.Windows.Media.Color]::FromRgb(137,180,250) }
            'error'      { [System.Windows.Media.Color]::FromRgb(243,139,168) }
        }
        $row.Indicator.Fill = [System.Windows.Media.SolidColorBrush]::new($color)
        if ($State -eq 'ok')                 { $row.Available = $true  }
        if ($State -in @('missing','error')) { $row.Available = $false }
        & $script:_SC_Refresh
    }

    $script:_SC_Check = {
        $allOk = $true

        foreach ($mod in $script:_SC_Modules) {
            $name = $mod.Name
            & $script:_SC_SetStatus $name 'checking' $script:T.SC_Checking
            $found = $false

            if ($script:_SC_IsOffline) {
                $localPath = Join-Path $script:_SC_DepsPath $name
                if (Test-Path $localPath) {
                    & $script:_SC_SetStatus $name 'ok' $script:T.SC_LocalCache
                    $found = $true
                }
            }

            if (-not $found) {
                $installed = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue
                if ($installed) {
                    $ver = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
                    & $script:_SC_SetStatus $name 'ok' ($script:T.SC_Installed -f $ver)
                    $found = $true
                }
            }

            if (-not $found) {
                & $script:_SC_SetStatus $name 'missing' $script:T.SC_Missing
                $allOk = $false
            }
        }

        if ($script:_SC_IsOffline) {
            $script:_SC_OfflineBanner.Visibility    = 'Visible'
            $script:_SC_OfflineBannerTitle.Text     = $script:T.SC_OfflineCacheTitle

            $missingLocally = $script:_SC_Modules | Where-Object {
                -not (Test-Path (Join-Path $script:_SC_DepsPath $_.Name))
            }

            if ($missingLocally) {
                $names = ($missingLocally | ForEach-Object { $_.Name }) -join ', '
                $script:_SC_OfflineBannerText.Text  = $script:T.SC_OfflineCacheMissing -f $names
                $script:_SC_BtnSaveDeps.Visibility  = 'Visible'
            }
            else {
                $script:_SC_OfflineBannerText.Text  = $script:T.SC_OfflineCacheOk
                $script:_SC_BtnSaveDeps.Visibility  = 'Collapsed'
            }
        }
        else {
            $script:_SC_OfflineBanner.Visibility    = 'Visible'
            $script:_SC_OfflineBannerTitle.Text     = $script:T.SC_OnlineTitle
            $script:_SC_OfflineBannerText.Text      = $script:T.SC_OnlineText
            $script:_SC_BtnSaveDeps.Visibility      = 'Visible'
        }

        if ($allOk) {
            $script:_SC_BtnLaunch.IsEnabled   = $true
            $script:_SC_BtnInstall.Visibility = 'Collapsed'
            $script:_SC_Status.Text = $script:T.SC_AllOk
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(166,227,161))
        }
        elseif ($script:_SC_IsOffline) {
            $script:_SC_BtnInstall.Visibility = 'Visible'
            $script:_SC_Status.Text = $script:T.SC_NeedInstall
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(249,168,37))
        }
        else {
            $script:_SC_BtnInstall.Visibility = 'Visible'
            $script:_SC_Status.Text = $script:T.SC_NeedInstall
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(249,168,37))
        }

        & $script:_SC_Refresh
    }

    # Preparation pour Install-Module (bouton "Installer les dependances")
    $script:_SC_PrepareGallery = {
        # TLS 1.2 requis par PSGallery (PS 5.1 utilise TLS 1.0 par defaut)
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        # NuGet requis pour Install-Module
        $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 `
            -Force -Scope CurrentUser -ErrorAction SilentlyContinue
    }

    $script:_SC_Install = {
        $script:_SC_BtnInstall.IsEnabled = $false
        $script:_SC_BtnQuit.IsEnabled    = $false
        $script:_SC_Progress.Visibility  = 'Visible'
        $script:_SC_Status.Foreground    = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))

        $anyError = $false
        $oldProgress    = $ProgressPreference
        $oldInformation = $InformationPreference
        $ProgressPreference    = 'SilentlyContinue'
        $InformationPreference = 'SilentlyContinue'

        $script:_SC_Status.Text = $script:T.SC_NuGet
        & $script:_SC_Refresh
        & $script:_SC_PrepareGallery

        foreach ($mod in $script:_SC_Modules) {
            $name = $mod.Name
            if ($script:_SC_DepRows[$name].Available) { continue }

            & $script:_SC_SetStatus $name 'installing' $script:T.SC_Installing
            $script:_SC_Status.Text = $script:T.SC_InstallingMod -f $name
            & $script:_SC_Refresh

            try {
                $null = Install-Module -Name $name -Repository PSGallery `
                    -Force -Scope CurrentUser `
                    -ErrorAction Stop -WarningAction SilentlyContinue
                $installed = Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue
                $ver = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
                & $script:_SC_SetStatus $name 'ok' ($script:T.SC_Installed -f $ver)
            }
            catch {
                & $script:_SC_SetStatus $name 'error' ($script:T.SC_ErrGeneric -f $_.Exception.Message)
                $anyError = $true
            }
        }

        $ProgressPreference    = $oldProgress
        $InformationPreference = $oldInformation
        $script:_SC_Progress.Visibility = 'Collapsed'
        $script:_SC_BtnQuit.IsEnabled   = $true

        if ($anyError) {
            $script:_SC_BtnInstall.IsEnabled = $true
            $script:_SC_Status.Text = $script:T.SC_InstallError
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
        }
        else {
            $script:_SC_BtnLaunch.IsEnabled   = $true
            $script:_SC_BtnInstall.Visibility = 'Collapsed'
            $script:_SC_Status.Text = $script:T.SC_InstallDone
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(166,227,161))
            $script:_SC_BtnSaveDeps.Visibility = 'Visible'
        }
        & $script:_SC_Refresh
    }

    $script:_SC_SaveDeps = {
        $confirm = [System.Windows.MessageBox]::Show(
            $script:T.SC_SaveConfirm, $script:T.SC_SaveTitle, 'YesNo', 'Question'
        )
        if ($confirm -ne 'Yes') { return }

        $script:_SC_BtnSaveDeps.IsEnabled = $false
        $script:_SC_BtnQuit.IsEnabled     = $false
        $script:_SC_Progress.Visibility   = 'Visible'
        $script:_SC_Status.Foreground     = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))

        $anyError = $false
        $oldProgress    = $ProgressPreference
        $oldInformation = $InformationPreference
        $ProgressPreference    = 'SilentlyContinue'
        $InformationPreference = 'SilentlyContinue'

        if (-not (Test-Path $script:_SC_DepsPath)) {
            New-Item -Path $script:_SC_DepsPath -ItemType Directory -Force | Out-Null
        }

        foreach ($mod in $script:_SC_Modules) {
            $name = $mod.Name
            & $script:_SC_SetStatus $name 'installing' $script:T.SC_Saving
            $script:_SC_Status.Text = $script:T.SC_DownloadingMod -f $name
            & $script:_SC_Refresh

            try {
                # TLS 1.2 obligatoire pour PSGallery
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                # Noms uniques pour eviter les conflits avec des runs precedents
                $runId       = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
                $localPath   = Join-Path $script:_SC_DepsPath $name
                $tempNupkg   = Join-Path $env:TEMP "$name.$runId.zip"
                $tempExtract = Join-Path $env:TEMP "$name.$runId.extract"

                # Nettoyage du dossier de destination
                if (Test-Path $localPath) { Remove-Item $localPath -Recurse -Force -ErrorAction SilentlyContinue }
                New-Item $localPath -ItemType Directory -Force | Out-Null

                # Telechargement ASYNCHRONE avec pompage du dispatcher : DownloadFile
                # synchrone figeait la fenetre pendant tout le telechargement
                # (~110 Mo pour MilestonePSTools). Meme idiome que Start-PumpingProcess.
                $wc = [System.Net.WebClient]::new()
                try {
                    $dlTask = $wc.DownloadFileTaskAsync("https://www.powershellgallery.com/api/v2/package/$name", $tempNupkg)
                    while (-not $dlTask.IsCompleted) {
                        $script:_SC_Win.Dispatcher.Invoke(
                            [System.Windows.Threading.DispatcherPriority]::Background, [Action]{})
                        Start-Sleep -Milliseconds 120
                    }
                    if ($dlTask.IsFaulted) {
                        $inner = $dlTask.Exception.InnerException
                        if ($inner) { throw $inner } else { throw $dlTask.Exception }
                    }
                } finally {
                    $wc.Dispose()
                }

                # Extraction (nupkg = ZIP) — API .NET directe, comme partout ailleurs
                # dans le projet (Initialize-Modules, Updater, .bat). Expand-Archive
                # re-liste chaque entree apres ecriture et echoue par intermittence
                # (ItemNotFoundException sur _rels\.rels) quand un antivirus retient
                # brievement les fichiers fraichement extraits.
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
                if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue }
                [System.IO.Compression.ZipFile]::ExtractToDirectory($tempNupkg, $tempExtract)

                # Copie des fichiers du module (sans les metadonnees NuGet)
                $excludeNames = @('[Content_Types].xml')
                $excludeExts  = @('.nuspec', '.psmdcp')
                $excludeDirs  = @('_rels', 'package')
                Get-ChildItem $tempExtract | Where-Object {
                    $_.Name -notin $excludeNames -and
                    $_.Extension -notin $excludeExts -and
                    $_.Name -notin $excludeDirs
                } | Copy-Item -Destination $localPath -Recurse -Force

                # Nettoyage
                Remove-Item $tempNupkg, $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

                & $script:_SC_SetStatus $name 'ok' $script:T.SC_CacheOk
            }
            catch {
                $errMsg = "$($_.Exception.GetType().Name): $($_.Exception.Message)"
                if ($_.Exception.InnerException) {
                    $errMsg += " | $($_.Exception.InnerException.Message)"
                }
                # MessageBox pour voir l'erreur exacte
                [System.Windows.MessageBox]::Show(
                    $errMsg, 'Erreur SaveDeps',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
                & $script:_SC_SetStatus $name 'error' ($script:T.SC_ErrGeneric -f $errMsg)
                $anyError = $true
            }
        }

        $ProgressPreference    = $oldProgress
        $InformationPreference = $oldInformation
        $script:_SC_Progress.Visibility   = 'Collapsed'
        $script:_SC_BtnSaveDeps.IsEnabled = $true
        $script:_SC_BtnQuit.IsEnabled     = $true

        if (-not $anyError) {
            $script:_SC_IsOffline = $true
            $script:_SC_BtnSaveDeps.Content = $script:T.SC_BtnUpdateCache
            & $script:_SC_Check
        }
        else {
            $script:_SC_Status.Text = $script:T.SC_SaveError
            $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
            & $script:_SC_Refresh
        }
    }

    $script:_SC_Win.Add_Loaded({
        try   { & $script:_SC_Check }
        catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.SC_ErrCheck -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
            ) | Out-Null
        }

        # Auto-installation si modules manquants et Internet accessible (test DNS rapide)
        $anyMissing = ($script:_SC_Modules | Where-Object { -not $script:_SC_DepRows[$_.Name].Available }).Count -gt 0
        if ($anyMissing) {
            $hasInternet = $false
            try { [System.Net.Dns]::GetHostEntry('www.powershellgallery.com') | Out-Null; $hasInternet = $true } catch {}

            if ($hasInternet) {
                $script:_SC_Status.Text = $script:T.SC_AutoInstalling
                $script:_SC_Status.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                    [System.Windows.Media.Color]::FromRgb(137,180,250))
                & $script:_SC_Refresh
                $script:_SC_Win.Dispatcher.BeginInvoke(
                    [System.Windows.Threading.DispatcherPriority]::ContextIdle,
                    [Action]{
                        try   { & $script:_SC_Install }
                        catch {
                            [System.Windows.MessageBox]::Show(
                                ($script:T.SC_ErrInstall -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
                            ) | Out-Null
                            $script:_SC_BtnInstall.IsEnabled = $true
                            $script:_SC_BtnQuit.IsEnabled    = $true
                            $script:_SC_Progress.Visibility  = 'Collapsed'
                        }
                    }
                ) | Out-Null
            }
        }

        # Verification de mise a jour differee (apres rendu complet de la fenetre)
        $script:_SC_Win.Dispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::ContextIdle,
            [Action]{ try { & $script:_SC_CheckAppUpdate } catch {} }
        ) | Out-Null
    })

    $script:_SC_BtnInstall.Add_Click({
        try   { & $script:_SC_Install }
        catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.SC_ErrInstall -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
            ) | Out-Null
            $script:_SC_BtnInstall.IsEnabled = $true
            $script:_SC_BtnQuit.IsEnabled    = $true
            $script:_SC_Progress.Visibility  = 'Collapsed'
        }
    })

    $script:_SC_BtnSaveDeps.Add_Click({
        try   { & $script:_SC_SaveDeps }
        catch {
            [System.Windows.MessageBox]::Show(
                ($script:T.SC_ErrGeneric -f $_), $script:T.SC_ErrTitle, 'OK', 'Error'
            ) | Out-Null
        }
    })

    $script:_SC_BtnLaunch.Add_Click({
        $script:_SC_Result = $true
        $script:_SC_Win.Close()
    })

    $script:_SC_BtnQuit.Add_Click({
        $script:_SC_Result = $false
        $script:_SC_Win.Close()
    })

    $script:_SC_AppUpdateBtn.Add_Click({
        $release = $script:_SC_AppRelease
        if (-not $release) { return }

        $confirm = [System.Windows.MessageBox]::Show(
            ($script:T.SC_AppUpdateConfirm -f $release.tag_name),
            $script:T.SC_AppUpdateTitle,
            'YesNo', 'Question'
        )
        if ($confirm -ne 'Yes') { return }

        $script:_SC_AppUpdateBtn.IsEnabled = $false
        $script:_SC_AppStatus.Visibility   = 'Visible'
        $script:_SC_AppStatus.Foreground   = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(166,173,200))
        $script:_SC_AppStatus.Text         = $script:T.SC_AppUpdating
        & $script:_SC_Refresh

        try {
            # Source de telechargement : on PRIVILEGIE le fichier .zip attache a la release
            # (comptabilise dans les telechargements GitHub, et verifie par SHA256). A defaut
            # (anciennes releases sans asset), on retombe sur l'archive source (zipball).
            $asset = $release.assets | Where-Object { $_.name -match '\.zip$' } | Select-Object -First 1
            if ($asset) {
                $zipUrl  = $asset.browser_download_url
                $isAsset = $true
            } else {
                $zipUrl  = $release.zipball_url
                $isAsset = $false
            }

            # SECURITE : l'URL doit pointer vers GitHub (HTTPS) et contenir le depot de confiance.
            # Couvre l'asset (/REPO/releases/download/...) et le zipball (/repos/REPO/zipball/...).
            $uri = [Uri]$zipUrl
            if ($uri.Scheme -ne 'https' -or
                $uri.Host -notmatch '(^|\.)github\.com$' -or
                $uri.AbsolutePath -notmatch [regex]::Escape($script:_SC_TrustedRepo)) {
                throw ($script:T.SC_AppUpdateUntrusted -f $zipUrl)
            }

            # SHA256 (dans les notes) = empreinte de l'ASSET. On la verifie uniquement quand on
            # telecharge l'asset ; en repli zipball, on s'appuie sur HTTPS + depot fige.
            $expectedHash = $null
            if ($isAsset -and $release.body -match '(?im)^\s*sha256\s*[:=]\s*([0-9a-fA-F]{64})\s*$') {
                $expectedHash = $Matches[1]
            }

            $tempDir = Join-Path $env:TEMP ("MilestoneToolkitUpdate_{0}" -f [guid]::NewGuid())
            $zipPath = Join-Path $tempDir 'release.zip'
            $extract = Join-Path $tempDir 'extract'
            New-Item -Path $tempDir,$extract -ItemType Directory -Force | Out-Null

            $dlHeaders = @{ 'User-Agent' = 'MilestoneToolkitUpdater' }
            Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath `
                -Headers $dlHeaders -UseBasicParsing -ErrorAction Stop

            if ($expectedHash) {
                $actualHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash
                if ($actualHash -ne $expectedHash) {
                    throw ($script:T.SC_AppUpdateBadHash -f $expectedHash, $actualHash)
                }
            }

            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extract)

            $srcRoot = Get-ChildItem -Path $extract | Where-Object PSIsContainer | Select-Object -First 1
            if (-not $srcRoot) { throw 'Archive vide ou structure inattendue.' }

            $updaterPath = Join-Path $script:_SC_AppRoot 'src\Core\Updater.ps1'
            $batPath     = Join-Path $script:_SC_AppRoot 'Demarrer Milestone Toolkit.bat'

            $ps = (Get-Command powershell -ErrorAction SilentlyContinue).Source
            if (-not $ps) { $ps = 'powershell.exe' }
            Start-Process -FilePath $ps `
                -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                              "`"$updaterPath`"",
                              "`"$($script:_SC_AppRoot)`"",
                              "`"$($srcRoot.FullName)`"",
                              "`"$batPath`""

            $script:_SC_Result = $false
            $script:_SC_Win.Close()
            exit 0
        }
        catch {
            $script:_SC_AppStatus.Text       = $script:T.SC_AppUpdateErr -f $_.Exception.Message
            $script:_SC_AppStatus.Foreground  = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.Color]::FromRgb(243,139,168))
            $script:_SC_AppStatus.Visibility  = 'Visible'
            $script:_SC_AppUpdateBtn.IsEnabled = $true
            & $script:_SC_Refresh
        }
    })

    [void]$script:_SC_Win.ShowDialog()
    return $script:_SC_Result
}
