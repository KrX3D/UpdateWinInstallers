param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "WinMerge"
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

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
