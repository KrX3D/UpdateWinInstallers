param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "WinSCP"
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

Write_LogEntry -Message "Beginne WinSCP Installation/Configuration" -Level "INFO"
Write-Host "WinSCP wird installiert" -foregroundcolor "magenta"

$installerPath = Join-Path $Serverip "Daten\Prog\WinSCP*.exe"
Write_LogEntry -Message "Suche nach Installer unter: $($installerPath)" -Level "DEBUG"

# Install WinSCP silently
$installer = Get-ChildItem -Path $installerPath | Select-Object -First 1
if ($installer) {
    Write_LogEntry -Message "Gefundene WinSCP Installer-Datei: $($installer.FullName)" -Level "INFO"
    [void](Invoke-InstallerFile -FilePath $installer.FullName -Arguments "/VERYSILENT", "/NORESTART", "/ALLUSERS" -Wait)
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

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
