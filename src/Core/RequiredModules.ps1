<#
.SYNOPSIS
    Liste canonique des modules PowerShell requis — SOURCE UNIQUE DE VERITE.
    Utilise par Initialize-Modules, Show-StartupCheck et Save-Dependencies.
    DescriptionKey : cle de traduction dans $script:T (voir src/Lang/*.ps1).
#>

function Get-RequiredModules {
    @(
        [pscustomobject]@{ Name = 'MilestonePSTools'; Required = $true; DescriptionKey = 'SC_ModuleDesc' }
        [pscustomobject]@{ Name = 'ImportExcel';      Required = $true; DescriptionKey = 'SC_ModuleExcelDesc' }
    )
}
