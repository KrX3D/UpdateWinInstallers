param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Notepad++"
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
        Write_LogEntry -Message "Konfigurationsdatei $($configPath) gefunden und importiert." -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

Write_LogEntry -Message "Notepad++ Installationsroutine gestartet." -Level "INFO"
Write-Host "Notepad++ wird installiert" -foregroundcolor "magenta"
$NotepadPlusPlusExeFiles = Get-ChildItem "$Serverip\Daten\Prog\npp*.exe" | Select-Object -First 1
Write_LogEntry -Message "Gefundene Notepad++ Installer (erste): $($($NotepadPlusPlusExeFiles.FullName))" -Level "DEBUG"

if ($NotepadPlusPlusExeFiles) {
    Write_LogEntry -Message "Starte Notepad++ Installer: $($NotepadPlusPlusExeFiles.FullName) mit Argument '/S' (synchronous - Wait)" -Level "INFO"
    Start-Process -Wait -FilePath $NotepadPlusPlusExeFiles.FullName -ArgumentList "/S"
    Write_LogEntry -Message "Notepad++ Installer ausgeführt: $($NotepadPlusPlusExeFiles.FullName)" -Level "SUCCESS"
}

$NotepadPlusPlusLnkFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Notepad++.lnk"
Write_LogEntry -Message "Prüfe Public Start-Menu Shortcut: $($NotepadPlusPlusLnkFile)" -Level "DEBUG"
if (Test-Path $NotepadPlusPlusLnkFile) {
    Write_LogEntry -Message "Entferne Startmenu Shortcut: $($NotepadPlusPlusLnkFile)" -Level "INFO"
	Remove-Item -Path $NotepadPlusPlusLnkFile -Force
    Write_LogEntry -Message "Startmenu Shortcut entfernt: $($NotepadPlusPlusLnkFile)" -Level "SUCCESS"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag ist gesetzt: stelle Notepad++ ein und kopiere Konfigurationen." -Level "INFO"

	$NotepadPlusPlusLnkFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Notepad++.lnk"
	$NotepadPlusPlusQuickLaunchPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    Write_LogEntry -Message "QuickLaunch Zielpfad: $($NotepadPlusPlusQuickLaunchPath)" -Level "DEBUG"

	if (Test-Path $NotepadPlusPlusLnkFile) {
		Write_LogEntry -Message "Verschiebe Startmenüeintrag $($NotepadPlusPlusLnkFile) nach $($NotepadPlusPlusQuickLaunchPath)" -Level "INFO"
		Move-Item -Path $NotepadPlusPlusLnkFile -Destination $NotepadPlusPlusQuickLaunchPath -Force
        Write_LogEntry -Message "Startmenüeintrag verschoben: Quelle=$($NotepadPlusPlusLnkFile) Ziel=$($NotepadPlusPlusQuickLaunchPath)" -Level "SUCCESS"
		#Remove-Item -Path $NotepadPlusPlusLnkFile -Force
	}

	#$NotepadRegFiles = Get-ChildItem "$Serverip\Daten\Prog\Notepad*.reg"
	#foreach ($regFile in $NotepadRegFiles) {
		#Start-Process -Wait -FilePath "REGEDIT.EXE" -ArgumentList "/S", $regFile.FullName
	#}
	
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
		-Path "$Serverip\Daten\Prog\Notepad_Taskbar.reg"
    Write_LogEntry -Message "RegistryImport Script aufgerufen: $($Serverip)\Daten\Customize_Windows\Scripte\RegistryImport.ps1 mit Path $($Serverip)\Daten\Prog\Notepad_Taskbar.reg" -Level "INFO"

	Write-Host "Notepad++ wird eingestellt." -foregroundcolor "Yellow"
    Write_LogEntry -Message "Starte Notepad++ Konfigurationsskript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Notepad_set_config.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\Notepad_set_config.ps1"
    Write_LogEntry -Message "Notepad++ Konfigurationsskript ausgeführt." -Level "SUCCESS"

	Start-Sleep -Seconds 5
    Write_LogEntry -Message "Beende explorer.exe (Stop-Process -Name explorer -Force) um Taskbar/Icons neu aufzubauen." -Level "INFO"
	Stop-Process -Name explorer -Force
    Write_LogEntry -Message "Explorer Prozess beendet." -Level "DEBUG"
	#Start-Process -FilePath explorer.exe

	$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
	Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"
	if (Test-Path $SetUserFTA) {
		Write_LogEntry -Message "Setze Dateizuordnungen für Notepad++ über SFTA: $($SetUserFTA)" -Level "INFO"
		Write-Host "	Notepad++ Dateizuordnung" -foregroundcolor "Yellow"
		$fileExtensions = @(".inf", ".ini", ".cs", ".css", ".pp", ".scp", ".wtx", ".gitignore", ".inc", ".sh", ".bsh", ".bash", ".cfg", ".fff", ".py", ".h", ".hh", ".hlp", ".hpp", ".htm", ".c", ".cpp", ".cxx", ".cc", ".m", ".md", ".mm", ".manifest", ".nfo", ".json", ".log", ".pem", ".php", ".php3", ".php4", ".php5", ".phps", ".ps1", ".txt", ".udf", ".UDL", ".udt", ".user", ".usr", ".uvu", ".wccf", ".xaml", ".xml", ".yaml", ".yml", ".vb", ".vbs")
		foreach ($extension in $fileExtensions) {
            Write_LogEntry -Message "Setze FTA: App='C:\Program Files\Notepad++\notepad++.exe' Extension=$($extension)" -Level "DEBUG"
			& $SetUserFTA --reg "C:\Program Files\Notepad++\notepad++.exe" $extension
		}
        Write_LogEntry -Message "Dateizuordnungen (FTAs) gesetzt für Notepad++." -Level "SUCCESS"
	}

	# Execute the .reg file
	#$regFile = "$Serverip\Daten\Prog\InstallationScripts\Notepad++.reg"
	#if (Test-Path $regFile) {
		#Write-Host "Importing registry file: $regFile"
		#Start-Process -FilePath "regedit.exe" -ArgumentList "/s", $regFile -Wait
	#}
	
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
		-Path "$Serverip\Daten\Prog\InstallationScripts\Notepad++.reg"
    Write_LogEntry -Message "RegistryImport für Notepad++ ausgeführt: $($Serverip)\Daten\Prog\InstallationScripts\Notepad++.reg" -Level "INFO"
}

Write_LogEntry -Message "Notepad++ Routine beendet." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
