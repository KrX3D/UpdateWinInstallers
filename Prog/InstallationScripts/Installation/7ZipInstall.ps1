param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "7-Zip"
$ScriptType = "Install"

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

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"

Write_LogEntry -Message "Prüfe Vorhandensein der Konfigurationsdatei: $($configPath)" -Level "DEBUG"
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

# Uninstall
$uninstallPath = "C:\Program Files\7-Zip\Uninstall.exe"
Write_LogEntry -Message "Uninstall-Pfad gesetzt: $($uninstallPath)" -Level "DEBUG"
if (Test-Path $uninstallPath) {
    Write_LogEntry -Message "7-Zip ist installiert, Deinstallation beginnt: $($uninstallPath)" -Level "INFO"
    Write-Host "7-Zip ist installiert, Deinstallation beginnt." -foregroundcolor "magenta"

    $uninstallArguments = "/S"
    Write_LogEntry -Message "Starte Deinstallation mit Argumenten: $($uninstallArguments)" -Level "DEBUG"

    Start-Process -FilePath $uninstallPath -ArgumentList $uninstallArguments -Wait

    Write_LogEntry -Message "7-Zip Deinstallation-Prozess beendet für: $($uninstallPath)" -Level "SUCCESS"
    Write-Host "	7-Zip wurde deinstalliert." -foregroundcolor "green"
    Start-Sleep -Seconds 3
} else {
    Write_LogEntry -Message "Kein Uninstall-Programm gefunden bei: $($uninstallPath)" -Level "DEBUG"
}

# Installation
Write_LogEntry -Message "Suche Installer auf Serverpfad: $($Serverip)\Daten\Prog\7z*.exe" -Level "DEBUG"
$installer = Get-ChildItem -Path "$Serverip\Daten\Prog\7z*.exe" | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue

if ($null -eq $installer -or $installer.Count -eq 0) {
    Write_LogEntry -Message "Keine Installer gefunden unter: $($Serverip)\Daten\Prog\7z*.exe" -Level "WARNING"
} else {
    $countInstallers = ($installer | Measure-Object).Count
    Write_LogEntry -Message "Gefundene Installer: $($countInstallers)" -Level "INFO"

	foreach ($exe in $installer) {
        Write_LogEntry -Message "Beginne Installation von: $($exe)" -Level "INFO"
	    Write-Host "7ZIP wird installiert: $exe" -foregroundcolor "magenta"
	    Start-Process -FilePath $exe -ArgumentList "/S" -Wait
        Write_LogEntry -Message "Installationsprozess beendet für: $($exe)" -Level "SUCCESS"
    }
}

$startMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\7-Zip"
Write_LogEntry -Message "Prüfe Startmenü-Pfad: $($startMenuPath)" -Level "DEBUG"
if (Test-Path $startMenuPath -PathType Container) {
    Write_LogEntry -Message "Startmenüeintrag gefunden, entferne: $($startMenuPath)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $startMenuPath -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($startMenuPath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag vorhanden bei: $($startMenuPath)" -Level "DEBUG"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag ist gesetzt; führe zusätzliche Installation-Schritte aus." -Level "INFO"

	$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
    Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"

	if (Test-Path $SetUserFTA -PathType Leaf) {
        Write_LogEntry -Message "SFTA gefunden. Setze Dateizuordnungen mit: $($SetUserFTA)" -Level "INFO"
		Write-Host "7Zip Dateizuordnung" -foregroundcolor "Yellow"
		$fileTypes = @(
			".001", ".7Z", ".zip", ".rar", ".cab", ".iso", ".img", ".xz", ".txz",
			".lzma", ".tar", ".cpio", ".bz2", ".bzip2", ".tbz2", ".tbz", ".gz",
			".gzip", ".tgz", ".tpz", ".z", ".taz", ".lzh", ".lha", ".rpm", ".deb",
			".arj", ".vhd", ".vhdx", ".wim", ".swm", ".esd", ".fat", ".ntfs",
			".dmg", ".hfs", ".xar", ".sqashfs", ".apfs"
		)
		$sortedFileTypes = $fileTypes | Sort-Object

        Write_LogEntry -Message "Anzahl Dateitypen für Zuordnung: $($sortedFileTypes.Count)" -Level "DEBUG"
		foreach ($type in $sortedFileTypes) {
            Write_LogEntry -Message "Setze Dateizuordnung für Typ: $($type) mit Programm: C:\Program Files\7-Zip\7zFM.exe" -Level "DEBUG"
			& $SetUserFTA --reg "C:\Program Files\7-Zip\7zFM.exe" $type
		}
        Write_LogEntry -Message "Dateizuordnungen mittels SFTA abgeschlossen." -Level "SUCCESS"
	} else {
        Write_LogEntry -Message "SFTA nicht gefunden unter: $($SetUserFTA)" -Level "WARNING"
	}

	# Execute the .reg file
	#$regFile = "$Serverip\Daten\Prog\InstallationScripts\7Zip.reg"
	#if (Test-Path $regFile) {
		#Write-Host "Importing registry file: $regFile"
		#Start-Process -FilePath "regedit.exe" -ArgumentList "/s", $regFile -Wait
	#}
    
    $registryImportScript = "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1"
    Write_LogEntry -Message "Starte Registry-Import Script: $($registryImportScript) mit PSHost: $($PSHostPath)" -Level "INFO"

	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
		-Path "$Serverip\Daten\Prog\InstallationScripts\7Zip.reg"

    Write_LogEntry -Message "Registry-Import Script aufgerufen: $($registryImportScript)" -Level "INFO"
} else {
    Write_LogEntry -Message "InstallationFlag nicht gesetzt; überspringe Benutzer-spezifische Einstellungen." -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
