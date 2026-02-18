param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Advanced Port Scanner"
$ScriptType  = "Install"

# === Logger-Header: automatisch eingefügt ===
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
    Finalize_LogSession
    exit
}

Write_LogEntry -Message "Beginne Installation: Advanced Port Scanner" -Level "INFO"
Write-Host "Advanced Port Scanner wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Suche Installer unter: $($InstallationFolder)\Advanced_Port_Scanner*.exe" -Level "DEBUG"
$InstallerPath = Get-ChildItem -Path "$InstallationFolder\Advanced_Port_Scanner*.exe" | Select-Object -ExpandProperty FullName
Write_LogEntry -Message "Gefundener Installer-Pfad: $($InstallerPath)" -Level "DEBUG"

Write_LogEntry -Message "Starte Installer: $($InstallerPath) mit Argumenten: /VERYSILENT /NORESTART /SP- /SUPPRESSMSGBOXS /NOCANCEL" -Level "INFO"
Start-Process -FilePath $InstallerPath -ArgumentList "/VERYSILENT", "/NORESTART", "/SP-", "/SUPPRESSMSGBOXS", "/NOCANCEL"
Write_LogEntry -Message "Start-Process für Advanced Port Scanner aufgerufen: $($InstallerPath)" -Level "INFO"

# Silent installer autostart app. Killing this behaviour using the code below
#https://community.chocolatey.org/packages/advanced-port-scanner#files
$t = 0
Write_LogEntry -Message "Warte auf Prozess 'advanced_port_scanner' (max. 20 Sekunden) - Beginne Polling-Schleife" -Level "DEBUG"
DO
{
	start-Sleep -Milliseconds 1000 #wait 100ms / loop
	$t++ #increase iteration 
} Until ($null -ne ($p=Get-Process -Name advanced_port_scanner* -ErrorAction SilentlyContinue) -or ($t -gt 20)) #wait until process is found or timeout reached

if($p) { #if process is found
    Write_LogEntry -Message "Process 'advanced_port_scanner' gefunden: PID $($p.Id) - beende Prozess" -Level "INFO"
	$p |Stop-Process  -Force #kill process 
	Write-Host "Beende Advanced Port Scanner Prozess"
    Write_LogEntry -Message "Process 'advanced_port_scanner' beendet: PID $($p.Id)" -Level "SUCCESS"
} else {
	Write-Host "Timeout für Advanced Port Scanner Prozess" #no process found but timeout reached
    Write_LogEntry -Message "Timeout erreicht; Prozess 'advanced_port_scanner' nicht gefunden nach $($t) Sekunden" -Level "WARNING"
}

# Close Advanced Port Scanner
#$ScriptPath = Get-ChildItem -Path "$InstallationFolder\AutoIt_Scripts\Advanced_Port_Scanner.exe" | Select-Object -ExpandProperty FullName
#Start-Process -FilePath $ScriptPath

# Remove Advanced Port Scanner Start Menu shortcut
start-Sleep -Milliseconds 3000
$startMenuAPS = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Advanced Port Scanner v2"
Write_LogEntry -Message "Prüfe Vorhandensein Startmenü-Eintrag: $($startMenuAPS)" -Level "DEBUG"
if (Test-Path $startMenuAPS) {
    Write_LogEntry -Message "Startmenüeintrag gefunden und wird entfernt: $($startMenuAPS)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $startMenuAPS -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($startMenuAPS)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden bei: $($startMenuAPS)" -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
