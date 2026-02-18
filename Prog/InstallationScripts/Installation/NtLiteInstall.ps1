param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "NTLite"
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
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

Write_LogEntry -Message "NTLite Installationsroutine gestartet." -Level "INFO"
Write-Host "NTLite wird installiert" -foregroundcolor "magenta"
# Install NTLite if it exists
$ntliteInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\NTLite\NTLite*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
Write_LogEntry -Message "Gefundener NTLite Installer: $($ntliteInstaller.FullName)" -Level "DEBUG"

if ($ntliteInstaller) {
    Write_LogEntry -Message "Starte NTLite Installer: $($ntliteInstaller.FullName) mit Argumenten '/VERYSILENT','/NORESTART' (synchronous - Wait)" -Level "INFO"
    Start-Process -FilePath $ntliteInstaller.FullName -ArgumentList '/VERYSILENT', '/NORESTART' -Wait
    Write_LogEntry -Message "NTLite Installer ausgeführt: $($ntliteInstaller.FullName)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein NTLite Installer gefunden unter: $($Serverip)\Daten\Prog\NTLite" -Level "WARNING"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag ist gesetzt: Kopiere Lizenz-/Einstellungen für NTLite falls vorhanden." -Level "INFO"

	# Copy NTLite license file if it exists
	$licenseFile = "$Serverip\Daten\Prog\NTLite\License\HP Desktop\license.dat"
	Write_LogEntry -Message "Prüfe Lizenzdatei: $($licenseFile)" -Level "DEBUG"
	if (Test-Path $licenseFile) {
		Write_LogEntry -Message "Konfigurations-Lizenzdatei gefunden: $($licenseFile). Kopiere nach 'C:\Program Files\NTLite\'" -Level "INFO"
		Copy-Item -Path $licenseFile -Destination 'C:\Program Files\NTLite\' -Force
        Write_LogEntry -Message "Lizenzdatei kopiert: Quelle=$($licenseFile) Ziel='C:\Program Files\NTLite\'" -Level "SUCCESS"
	} else {
        Write_LogEntry -Message "Lizenzdatei nicht gefunden: $($licenseFile)" -Level "DEBUG"
    }

	# Copy NTLite settings file if it exists
	$settingsFile = "$Serverip\Daten\Prog\NTLite\License\HP Desktop\settings.xml"
	Write_LogEntry -Message "Prüfe Settings-Datei: $($settingsFile)" -Level "DEBUG"
	if (Test-Path $settingsFile) {
		Write_LogEntry -Message "Einstellungsdatei gefunden: $($settingsFile). Kopiere nach 'C:\Program Files\NTLite\'" -Level "INFO"
		Copy-Item -Path $settingsFile -Destination 'C:\Program Files\NTLite\' -Force
        Write_LogEntry -Message "Settings-Datei kopiert: Quelle=$($settingsFile) Ziel='C:\Program Files\NTLite\'" -Level "SUCCESS"
	} else {
        Write_LogEntry -Message "Settings-Datei nicht gefunden: $($settingsFile)" -Level "DEBUG"
    }
}

# Remove NTLite Start Menu shortcuts if they exist
$ntliteStartMenuPath = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\NTLite'
Write_LogEntry -Message "Prüfe Startmenüpfad: $($ntliteStartMenuPath)" -Level "DEBUG"
if (Test-Path $ntliteStartMenuPath) {
	Write_LogEntry -Message "Startmenüeintrag gefunden: $($ntliteStartMenuPath). Entferne..." -Level "INFO"
    Remove-Item -Path $ntliteStartMenuPath -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($ntliteStartMenuPath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden unter: $($ntliteStartMenuPath)" -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
