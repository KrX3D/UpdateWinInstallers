param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "WinMerge"
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

#Bei Update wird ohne deinstallation die neue Version richtig installiert

Write_LogEntry -Message "Beginne Installation/Konfiguration von WinMerge" -Level "INFO"
Write-Host "WinMerge wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Suche nach WinMerge Installer unter: $($Serverip)\Daten\Prog\WinMerge*.exe" -Level "DEBUG"
$winMergeExeFile = Get-ChildItem "$Serverip\Daten\Prog\WinMerge*.exe" | Select-Object -First 1
if ($winMergeExeFile) {
    Write_LogEntry -Message "Gefundene WinMerge-Installer-Datei: $($winMergeExeFile.FullName)" -Level "INFO"
    Start-Process -Wait $winMergeExeFile.FullName -ArgumentList "/VERYSILENT", "/NORESTART"
    Write_LogEntry -Message "Start-Process ausgeführt für WinMerge Installer: $($winMergeExeFile.FullName)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein WinMerge-Installer gefunden unter: $($Serverip)\Daten\Prog\WinMerge*.exe" -Level "WARNING"
}

$winMergeShortcut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\WinMerge\WinMerge.lnk"
Write_LogEntry -Message "Erwarteter Shortcut-Pfad: $($winMergeShortcut)" -Level "DEBUG"
$publicLink = Join-Path ([Environment]::GetFolderPath('CommonDesktopDirectory')) 'WinMerge.lnk'
$desktopFolder = [Environment]::GetFolderPath("Desktop")
$desktopShortcut = Join-Path $desktopFolder "WinMerge.lnk"

if (Test-Path $winMergeShortcut) {
    Write_LogEntry -Message "WinMerge Startmenüshortcut gefunden: $($winMergeShortcut). Verschiebe nach: $($desktopShortcut)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird verschoben." -foregroundcolor "Cyan"
    Move-Item -Path $winMergeShortcut -Destination $desktopShortcut -Force
    Write_LogEntry -Message "Shortcut verschoben: $($desktopShortcut)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüshortcut gefunden unter: $($winMergeShortcut)" -Level "DEBUG"
}

if ( (Test-Path $desktopShortcut -PathType Leaf) -and (Test-Path $publicLink -PathType Leaf) ) {
    Write_LogEntry -Message "Doppelte Verknüpfung gefunden: Desktop $($desktopShortcut) und Public $($publicLink). Entferne Public Link." -Level "INFO"
    Write-Host "Doppelte Verknüpfung wird vom Desktop entfernt:" $publicLink -foregroundcolor "Cyan"
    Remove-Item $publicLink
    Write_LogEntry -Message "Entfernt: $($publicLink)" -Level "SUCCESS"
}

$winMergeMenuFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\WinMerge"
Write_LogEntry -Message "Prüfe Startmenüordner: $($winMergeMenuFolder)" -Level "DEBUG"
if (Test-Path $winMergeMenuFolder) {
    Write_LogEntry -Message "WinMerge Startmenüordner gefunden: $($winMergeMenuFolder). Entferne Ordner." -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $winMergeMenuFolder -Recurse -Force
    Write_LogEntry -Message "Entfernt Startmenüordner: $($winMergeMenuFolder)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein WinMerge Startmenüordner gefunden unter: $($winMergeMenuFolder)" -Level "DEBUG"
}

Write_LogEntry -Message "Starte Registry-Import für WinMerge Kontextmenü-Entfernung: $($Serverip)\Daten\Customize_Windows\Reg\Kontextmenu\Winmerge_kontextmenu_entfernen.reg" -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
	-Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Winmerge_kontextmenu_entfernen.reg"
Write_LogEntry -Message "Registry-Import aufgerufen für: $($Serverip)\Daten\Customize_Windows\Reg\Kontextmenu\Winmerge_kontextmenu_entfernen.reg" -Level "DEBUG"

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag ist $($InstallationFlag) -> führe zusätzliche Konfigurationen für WinMerge aus." -Level "INFO"
	Write-Host "WinMerge wird kofiguriert"
	Write_LogEntry -Message "Starte Registry-Import für WinMerge Config: $($Serverip)\Daten\Prog\WinMerge_Config.reg" -Level "INFO"

	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
		-Path "$Serverip\Daten\Prog\WinMerge_Config.reg"

    Write_LogEntry -Message "Registry-Import aufgerufen für: $($Serverip)\Daten\Prog\WinMerge_Config.reg" -Level "DEBUG"
} else {
    Write_LogEntry -Message "InstallationFlag ist $($InstallationFlag) -> keine zusätzlichen WinMerge-Konfigurationen ausgeführt." -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
