param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Arduino"
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

#Bei Update wird ohne deinstallation die neue Version richtig installiert

$ArduinoDocumentFolder = "$Serverip\Daten\Prog\Arduino\Arduino"
$Arduino15Folder = "$Serverip\Daten\Projekte\ESP\Arduino15"

$ArduinoDocumentDestinationFolder = "$env:USERPROFILE\Documents\Arduino"
$Arduino15DestinationFolder = "$env:USERPROFILE\AppData\Local\Arduino15"

Write_LogEntry -Message "Arduino-Quellordner: $($ArduinoDocumentFolder); Arduino15-Quellordner: $($Arduino15Folder)" -Level "DEBUG"
Write_LogEntry -Message "Zielordner: Documents: $($ArduinoDocumentDestinationFolder); Arduino15: $($Arduino15DestinationFolder)" -Level "DEBUG"

#@echo  Arduino Zertifikate werden installiert
#FOR %%A IN ("\\%Serverip%\Daten\Prog\Arduino\*.cer") DO certutil -f -addstore "TrustedPublisher" %%A
#@echo  ##############################
#@echo  Arduino wird installiert
#FOR %%A IN ("\\%Serverip%\Daten\Prog\Arduino\arduino-*_old.exe") DO start /wait %%A /S

Write_LogEntry -Message "Suche Arduino-Installer unter: $($Serverip)\Daten\Prog\Arduino\arduino-ide*.exe" -Level "DEBUG"
Write-Host "Arduino wird installiert" -foregroundcolor "magenta"
$arduinoExe = Get-ChildItem "$Serverip\Daten\Prog\Arduino\arduino-ide*.exe" | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue
if ($arduinoExe) {
    Write_LogEntry -Message "Gefundener Arduino-Installer: $($arduinoExe)" -Level "INFO"
    Write_LogEntry -Message "Starte Arduino Installer: $($arduinoExe) mit Argument: /S" -Level "INFO"
    [void](Invoke-InstallerFile -FilePath $arduinoExe -Arguments "/S" -Wait)
    Write_LogEntry -Message "Arduino Installer beendet: $($arduinoExe)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Arduino-Installer gefunden unter: $($Serverip)\Daten\Prog\Arduino\arduino-ide*.exe" -Level "WARNING"
}

if ((Test-Path $ArduinoDocumentFolder) -and ($InstallationFlag -eq $true)) {
    Write_LogEntry -Message "ArduinoDocumentFolder vorhanden und InstallationFlag true: $($ArduinoDocumentFolder)" -Level "INFO"
    if (-not (Test-Path $ArduinoDocumentDestinationFolder)) {
        Write_LogEntry -Message "Ziel Documents-Ordner existiert nicht: $($ArduinoDocumentDestinationFolder). Starte robocopy." -Level "INFO"
        Write-Host "	Arduino Bibliotheken werden wiederhergestellt" -foregroundcolor "Cyan"
        #Copy-Item $ArduinoDocumentFolder $ArduinoDocumentDestinationFolder -Recurse -Force
        robocopy $ArduinoDocumentFolder $ArduinoDocumentDestinationFolder /E /Z /NP /R:1 /W:1 /nfl /MT:64
        Write_LogEntry -Message "Robocopy-Aufruf abgeschlossen: Quelle=$($ArduinoDocumentFolder), Ziel=$($ArduinoDocumentDestinationFolder)" -Level "SUCCESS"
    }
	else
	{
        Write_LogEntry -Message "Arduino Bibliotheken werden nicht wiederhergestellt, Ziel existiert bereits: $($ArduinoDocumentDestinationFolder)" -Level "DEBUG"
		Write-Host "Arduino Bibliotheken werden nicht wiederhergestellt, da sie bereits vorhanden sind." -foregroundcolor "Yellow"
	}
}

if ((Test-Path $Arduino15Folder) -and ($InstallationFlag -eq $true)) {
    Write_LogEntry -Message "Arduino15Folder vorhanden und InstallationFlag true: $($Arduino15Folder)" -Level "INFO"
    if (-not (Test-Path $Arduino15DestinationFolder)) {
        Write_LogEntry -Message "Ziel Arduino15-Ordner existiert nicht: $($Arduino15DestinationFolder). Starte robocopy." -Level "INFO"
		Write-Host "	Arduino Boards werden wiederhergestellt" -foregroundcolor "Cyan"
        #Copy-Item $Arduino15Folder $Arduino15DestinationFolder -Recurse -Force
        robocopy $Arduino15Folder $Arduino15DestinationFolder /E /Z /NP /R:1 /W:1 /nfl /MT:64
        Write_LogEntry -Message "Robocopy-Aufruf abgeschlossen: Quelle=$($Arduino15Folder), Ziel=$($Arduino15DestinationFolder)" -Level "SUCCESS"
    }
	else
	{
        Write_LogEntry -Message "Arduino Boards werden nicht wiederhergestellt, Ziel existiert bereits: $($Arduino15DestinationFolder)" -Level "DEBUG"
		Write-Host "Arduino Boards werden nicht wiederhergestellt, da sie bereits vorhanden sind." -foregroundcolor "Yellow"
	}
}

if ($InstallationFlag -eq $true) {
	$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
    Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"

	if (Test-Path $SetUserFTA -PathType Leaf) {
        Write_LogEntry -Message "SFTA gefunden: $($SetUserFTA) - beginne Dateizuordnungen" -Level "INFO"
		Write-Host "Arduino Dateizuordnung" -foregroundcolor "Yellow"
		$fileTypes = @(
			".ino"
		)
		$sortedFileTypes = $fileTypes | Sort-Object

        Write_LogEntry -Message "Anzahl Dateitypen für Zuordnung: $($sortedFileTypes.Count)" -Level "DEBUG"
		foreach ($type in $sortedFileTypes) {
            Write_LogEntry -Message "Setze Dateizuordnung für Typ: $($type) mit Programm: $($env:USERPROFILE)\AppData\Local\Programs\Arduino IDE\Arduino IDE.exe" -Level "DEBUG"
			& $SetUserFTA --reg "$env:USERPROFILE\AppData\Local\Programs\Arduino IDE\Arduino IDE.exe" $type
            Write_LogEntry -Message "Aufruf SFTA beendet für Typ: $($type)" -Level "DEBUG"
		}
        Write_LogEntry -Message "Alle Dateizuordnungen mittels SFTA abgeschlossen." -Level "SUCCESS"
	} else {
        Write_LogEntry -Message "SFTA nicht gefunden unter: $($SetUserFTA)" -Level "WARNING"
	}
}

$shortcutArduino = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Arduino.lnk"
$shortcutArduinoIDE = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Arduino IDE.lnk"

Write_LogEntry -Message "Prüfe Shortcut: $($shortcutArduino)" -Level "DEBUG"
if (Test-Path $shortcutArduino) {
    Write_LogEntry -Message "Startmenüeintrag gefunden und wird entfernt: $($shortcutArduino)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $shortcutArduino -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($shortcutArduino)" -Level "SUCCESS"
}

Write_LogEntry -Message "Prüfe Shortcut: $($shortcutArduinoIDE)" -Level "DEBUG"
if (Test-Path $shortcutArduinoIDE) {
    Write_LogEntry -Message "Startmenüeintrag gefunden und wird entfernt: $($shortcutArduinoIDE)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $shortcutArduinoIDE -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($shortcutArduinoIDE)" -Level "SUCCESS"
}

#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Arduino_kontextmenu_entfernen.reg" -Wait
$registryImportScript = "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1"
Write_LogEntry -Message "Rufe RegistryImport Script auf: $($registryImportScript)" -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
	-Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Arduino_kontextmenu_entfernen.reg"
Write_LogEntry -Message "RegistryImport Script aufgerufen: $($registryImportScript)" -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
