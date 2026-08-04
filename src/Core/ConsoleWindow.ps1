<#
.SYNOPSIS
    Masquage/affichage de la fenetre console — helper partage (Bootstrap + App).
    Evite la duplication du P/Invoke ConsoleHider/ConsoleHelper.
#>

if (-not ('MT.ConsoleWindow' -as [type])) {
    Add-Type -Namespace 'MT' -Name 'ConsoleWindow' -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@ -ErrorAction SilentlyContinue
}

function Hide-Console {
    try { [MT.ConsoleWindow]::ShowWindow([MT.ConsoleWindow]::GetConsoleWindow(), 0) | Out-Null } catch {}
}

function Show-Console {
    try { [MT.ConsoleWindow]::ShowWindow([MT.ConsoleWindow]::GetConsoleWindow(), 5) | Out-Null } catch {}
}
