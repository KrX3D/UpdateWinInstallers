param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "PowerShell 7"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

#Bei Update wird ohne deinstallation die neue Version richtig installiert

Write_LogEntry -Message "PowerShell Installationsroutine gestartet." -Level "INFO"
Write-Host "PowerShell wird installiert" -foregroundcolor "magenta"

# Install PowerShell 7 if it exists
$powerShellInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\PowerShell-*-win-x64.msi" -ErrorAction SilentlyContinue | Select-Object -First 1
Write_LogEntry -Message ("Gefundener PowerShell Installer: " + $(if ($powerShellInstaller) { $powerShellInstaller.FullName } else { '<none>' })) -Level "DEBUG"

if ($powerShellInstaller) {
    try {
        Write_LogEntry -Message "Starte PowerShell Installer: $($(if ($powerShellInstaller) { $powerShellInstaller.FullName } else { '<none>' })) mit stillen Parametern." -Level "INFO"
        [void](Invoke-InstallerFile -FilePath $powerShellInstaller.FullName -Arguments 'REGISTER_MANIFEST=0', 'ENABLE_PSREMOTING=1', 'ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1', 'ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1', 'ADD_PATH=1', 'DISABLE_TELEMETRY=1', 'ENABLE_MU=0', 'USE_MU=0', '/qn' -Wait)
        Write_LogEntry -Message "PowerShell Installer ausgeführt: $($powerShellInstaller.FullName)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Ausführen des PowerShell Installers $($(if ($powerShellInstaller) { $powerShellInstaller.FullName } else { '<none>' })): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein PowerShell Installer gefunden unter $($Serverip)\Daten\Prog" -Level "WARNING"
}

# Move PowerShell 7 Start Menu shortcut if it exists
$powerShellShortcut = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerShell\PowerShell 7 (x64).lnk'
Write_LogEntry -Message "Prüfe Shortcut-Pfad: $($powerShellShortcut)" -Level "DEBUG"
if (Test-Path $powerShellShortcut) {
    try {
        Write_LogEntry -Message "Verschiebe PowerShell Shortcut: $($powerShellShortcut) -> 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs'" -Level "INFO"
        Move-Item -Path $powerShellShortcut -Destination 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs' -Force
        Write_LogEntry -Message "Shortcut verschoben: $($powerShellShortcut)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Verschieben des Shortcuts $($powerShellShortcut): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "PowerShell Shortcut nicht gefunden: $($powerShellShortcut)" -Level "DEBUG"
}

# Remove PowerShell Start Menu shortcuts if they exist
$powerShellStartMenuPath = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerShell'
Write_LogEntry -Message "Prüfe Startmenüordner: $($powerShellStartMenuPath)" -Level "DEBUG"
if (Test-Path $powerShellStartMenuPath) {
    try {
        Write_LogEntry -Message "Entferne Startmenüeintrag: $($powerShellStartMenuPath)" -Level "INFO"
        Remove-Item -Path $powerShellStartMenuPath -Recurse -Force
        Write_LogEntry -Message "Startmenüeintrag entfernt: $($powerShellStartMenuPath)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Entfernen des Startmenüeintrags $($powerShellStartMenuPath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein PowerShell Startmenüeintrag gefunden unter: $($powerShellStartMenuPath)" -Level "DEBUG"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
