param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Git"
$ScriptType  = "Install"

# === Logger-Header: automatisch eingefgt ===
$parentPath  = Split-Path -Path $PSScriptRoot -Parent
$modulePath  = Join-Path -Path $parentPath -ChildPath 'Modules\Logger\Logger.psm1'

if (Test-Path $modulePath) {
    Import-Module -Name $modulePath -Force -ErrorAction Stop

    if (-not (Get-Variable -Name logRoot -Scope Script -ErrorAction SilentlyContinue)) {
        $logRoot = Join-Path -Path $parentPath -ChildPath 'Log'
    }
    Set_LoggerConfig -LogRootPath $logRoot | Out-Null

    if (Get-Command -Name Initialize_LogSession -ErrorAction SilentlyContinue) {
        Initialize_LogSession -ProgramName $ProgramName -ScriptType $ScriptType | Out-Null #-WriteSystemInfo
    }
}
# === Ende Logger-Header ===

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt: $($ProgramName); ScriptType gesetzt: $($ScriptType)" -Level "DEBUG"

# DeployToolkit helpers
$dtPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Modules\DeployToolkit\DeployToolkit.psm1"
if (Test-Path $dtPath) {
    Import-Module -Name $dtPath -Force -ErrorAction Stop
} else {
    if (Get-Command -Name Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "DeployToolkit nicht gefunden: $dtPath" -Level "WARNING"
    } else {
        Write-Warning "DeployToolkit nicht gefunden: $dtPath"
    }
}

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Konfigurationspfad gesetzt: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei $($configPath) gefunden und importiert." -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Script beendet wegen fehlender Konfigurationsdatei: $($configPath)" -Level "ERROR"
    pause
    exit
}

Write_LogEntry -Message "Starte Git-Installation (falls vorhanden) mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write-Host "Git wird installiert" -foregroundcolor "magenta"

# Install Git
$gitInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\Git*.exe" | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue
if ($gitInstaller) {
    Write_LogEntry -Message "Git-Installer gefunden: $($gitInstaller)" -Level "DEBUG"
    Write_LogEntry -Message "Starte Git-Installer: $($gitInstaller) mit stillen Parametern (Wait)" -Level "INFO"
    Start-Process -FilePath $gitInstaller -ArgumentList '/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXS', '/NOCANCEL', '/NORESTART', '/NOICONS', '/PathOption=CmdTools' -Wait
    Write_LogEntry -Message "Git-Installer beendet: $($gitInstaller)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Git-Installer gefunden unter: $($Serverip)\Daten\Prog (Pattern: Git*.exe)" -Level "WARNING"
    Write-Host "Kein Git-Installer gefunden: $Serverip\Daten\Prog\Git*.exe" -ForegroundColor "Yellow"
}

#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Git_Kontextmenu_entfernen.reg" -Wait
Write_LogEntry -Message "Starte Registry-Import Skript fr Git Kontextmenu (via PSHostPath): $($PSHostPath) mit Pfad: $($Serverip)\Daten\Customize_Windows\Reg\Kontextmenu\Git_Kontextmenu_entfernen.reg" -Level "INFO"
try {
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
        -Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Git_Kontextmenu_entfernen.reg"
    Write_LogEntry -Message "Registry-Importskript aufgerufen: $($Serverip)\Daten\Customize_Windows\Scripte\RegistryImport.ps1" -Level "SUCCESS"
} catch {
    Write_LogEntry -Message "Fehler beim Aufruf des Registry-Importskripts: $($_)" -Level "ERROR"
    Write-Host "Fehler beim Ausfhren des Registry-Importskripts: $($_)" -ForegroundColor "Red"
}

# === Logger-Footer: automatisch eingefgt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
