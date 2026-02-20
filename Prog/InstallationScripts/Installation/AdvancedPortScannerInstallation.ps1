param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Advanced Port Scanner"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt: $($ProgramName); ScriptType gesetzt: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

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

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
