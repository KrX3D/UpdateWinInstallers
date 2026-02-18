param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "WinSCP"
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
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Konfigurationspfad gesetzt: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei gefunden und importiert: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

Write_LogEntry -Message "Beginne WinSCP Installation/Configuration" -Level "INFO"
Write-Host "WinSCP wird installiert" -foregroundcolor "magenta"

$installerPath = Join-Path $Serverip "Daten\Prog\WinSCP*.exe"
Write_LogEntry -Message "Suche nach Installer unter: $($installerPath)" -Level "DEBUG"

# Install WinSCP silently
$installer = Get-ChildItem -Path $installerPath | Select-Object -First 1
if ($installer) {
    Write_LogEntry -Message "Gefundene WinSCP Installer-Datei: $($installer.FullName)" -Level "INFO"
    Start-Process -FilePath $installer.FullName -ArgumentList "/VERYSILENT", "/NORESTART", "/ALLUSERS" -Wait
    Write_LogEntry -Message "Start-Process ausgeführt für WinSCP Installer: $($installer.FullName)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein WinSCP-Installer gefunden unter: $($installerPath)" -Level "WARNING"
}

# Delete the Start Menu shortcut
$shortcutPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\WinSCP.lnk"
Write_LogEntry -Message "Erwarte Shortcut-Pfad: $($shortcutPath)" -Level "DEBUG"
if (Test-Path $shortcutPath) {
    Write_LogEntry -Message "Shortcut gefunden und wird entfernt: $($shortcutPath)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $shortcutPath -Force
    Write_LogEntry -Message "Entfernt Shortcut: $($shortcutPath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Shortcut gefunden unter: $($shortcutPath)" -Level "DEBUG"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag ist $($InstallationFlag) -> Konfigurationsschritte werden ausgeführt." -Level "INFO"
	Write-Host "WinSCP wird kofiguriert"
	Write-Host "	SSH Host Keys werden gesetzt"
	Write_LogEntry -Message "Starte Ermittlung PCName via Script" -Level "DEBUG"

	#PC Name wird gesucht
	$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
	Write_LogEntry -Message "Script zur PC-Ermittlung: $($scriptPath)" -Level "DEBUG"

	try {
		Write_LogEntry -Message "Aufruf des Scripts zur PC-Ermittlung: $($scriptPath)" -Level "INFO"
		#$PCName = & $scriptPath
		#$PCName = & "powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
		$PCName = & $PSHostPath `
			-NoLogo -NoProfile -ExecutionPolicy Bypass `
			-File $scriptPath `
			-Verbose:$false
		Write_LogEntry -Message "PCName erfolgreich ermittelt: $($PCName)" -Level "SUCCESS"
	} catch {
		Write_LogEntry -Message "Fehler beim Ausführen des Scripts $($scriptPath): $($_)" -Level "ERROR"
		Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
		Pause
		Exit
	}
	
	# Define the backup file path
	$backupFilePath = "$Serverip\Daten\Windows_Backup\$PCName\WinSCP_Backup.reg"
	Write_LogEntry -Message "Backup-Pfad für WinSCP gesetzt: $($backupFilePath)" -Level "DEBUG"

	# Check if the backup file exists
	if (Test-Path $backupFilePath) {
		Write_LogEntry -Message "Backup-Datei gefunden: $($backupFilePath). Import starte." -Level "INFO"
		Write-Host ""
		Write-Host "Einstellungen und Connections wiederherstellen." -ForegroundColor "Yellow"
		try {
			reg import $backupFilePath
			Write_LogEntry -Message "Registry-Import erfolgreich für: $($backupFilePath)" -Level "SUCCESS"
		} catch {
			Write_LogEntry -Message "Fehler beim Registry-Import $($backupFilePath): $($_)" -Level "ERROR"
		}
	} else {
		Write_LogEntry -Message "Kein Backup gefunden unter: $($backupFilePath)" -Level "WARNING"
		#Write-Host "Registry file not found: $backupFilePath"
	}
} else {
    Write_LogEntry -Message "InstallationFlag ist $($InstallationFlag) -> keine Konfiguration durchgeführt." -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
