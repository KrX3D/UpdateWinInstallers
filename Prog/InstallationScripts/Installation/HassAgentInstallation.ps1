param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Hass.Agent"
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
Write_LogEntry -Message "Programm: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

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
    exit
}

Write_LogEntry -Message "Beginne Abschnitt: Hass Agent Installation/Update" -Level "INFO"
Write-Host "Hass Agent wird installiert" -foregroundcolor "magenta"

# Function to close HASS Agent if it is running
function Close-HassAgent {
    Write_LogEntry -Message "Close-HassAgent: Aufruf" -Level "DEBUG"
    $hassAgentProcess = Get-Process -Name "HASS.Agent" -ErrorAction SilentlyContinue
    if ($hassAgentProcess) {
        Write_LogEntry -Message "Close-HassAgent: Gefundene Prozesse: $($($hassAgentProcess).Count)" -Level "INFO"
        Write-Host "Hass Agent wird beendet..." -foregroundcolor "Yellow"
        $hassAgentProcess | Stop-Process -Force
        Write_LogEntry -Message "Close-HassAgent: Prozesse gestoppt." -Level "SUCCESS"
        Start-Sleep -Seconds 2 # Give it a moment to close
    } else {
        Write_LogEntry -Message "Close-HassAgent: Kein laufender HASS.Agent Prozess gefunden." -Level "DEBUG"
    }
}

# Function to start HASS Agent if it is not running
function Start-HassAgent {
    Write_LogEntry -Message "Start-HassAgent: Aufruf" -Level "DEBUG"
    $hassAgentProcess = Get-Process -Name "HASS.Agent" -ErrorAction SilentlyContinue
    if ($hassAgentProcess) {
        Write_LogEntry -Message "Start-HassAgent: HASS.Agent bereits laufend. Count=$($($hassAgentProcess).Count)" -Level "INFO"
        Write-Host "Hass Agent is already running." -ForegroundColor "Green"
    } else {
        # Get the current user's AppData local path
        $AppDataLocal = [System.Environment]::GetFolderPath("LocalApplicationData")
        $programPath = Join-Path $AppDataLocal "HASS.Agent\Client\HASS.Agent.exe"
        Write_LogEntry -Message "Start-HassAgent: Programmpfad ermittelt: $($programPath)" -Level "DEBUG"

        # Check if the file exists and start HASS.Agent
        if (Test-Path $programPath) {
            try {
                Start-Process $programPath
                Write_LogEntry -Message "Start-HassAgent: HASS.Agent gestartet: $($programPath)" -Level "SUCCESS"
                Write-Host "Hass Agent started successfully." -ForegroundColor "Green"
            } catch {
                Write_LogEntry -Message "Start-HassAgent: Fehler beim Starten von $($programPath): $($_)" -Level "ERROR"
                Write-Host "Failed to start HASS.Agent. Error: $_" -ForegroundColor "Red"
            }
        } else {
            Write_LogEntry -Message "Start-HassAgent: Programmpfad existiert nicht: $($programPath)" -Level "WARNING"
            Write-Host "The specified path does not exist: $programPath" -ForegroundColor "Yellow"
        }
    }
}

$agentPath = "$Serverip\Daten\Projekte\Smart_Home\HASS_Agent\HASS.Agent*.exe"
Write_LogEntry -Message "Agent Path Wildcard: $($agentPath)" -Level "DEBUG"
#$agentPath = "$Serverip\Daten\Projekte\Smart_Home\HASS_Agent\HASS.Agent.Installer.exe"
$agentFile = Get-ChildItem -Path $agentPath | Select-Object -First 1 -ExpandProperty FullName
Write_LogEntry -Message "Gefundene Agent-Datei (erste): $($agentFile)" -Level "DEBUG"

if (Test-Path $agentFile) {
    Write_LogEntry -Message "Agent-Datei vorhanden: $($agentFile). Vorbereitung zum Stop/Install." -Level "INFO"
	# Close HASS Agent if running
	Close-HassAgent
    Write_LogEntry -Message "Nach Close-HassAgent: fortfahren mit Installation" -Level "DEBUG"
	
    # Start the AutoIt script to handle the message box
    #$autoItExePath = "$Serverip\Daten\Prog\AutoIt_Scripts\HassAgentPressJa.exe" # Change this if your AutoIt executable is in a different location
    #Write_LogEntry -Message "AutoIt Exe Pfad: $($autoItExePath)" -Level "DEBUG"
    #Start-Process -FilePath $autoItExePath
    #Write_LogEntry -Message "AutoIt Helper gestartet: $($autoItExePath)" -Level "INFO"

	#Start-Process -FilePath $agentFile -ArgumentList "/exenoui /exenoupdates /passive /norestart" -Wait
    Write_LogEntry -Message "Starte Agent-Installer: $($agentFile) mit stillen Parametern (Wait)" -Level "INFO"
	Start-Process -FilePath $agentFile -ArgumentList "/SP- /SILENT /SUPPRESSMESGBOXES /NORESTART" -Wait
    Write_LogEntry -Message "Agent-Installer beendet: $($agentFile)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Agent-Datei nicht gefunden: $($agentPath)" -Level "ERROR"
	Write-Host "Hass Agent Datei nicht gefunden." -foregroundcolor "Red"
}

$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
Write_LogEntry -Message "Externes Script Pfad für PCName: $($scriptPath)" -Level "DEBUG"

try {
	#$PCName = & $scriptPath
	#$PCName = & "powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
	$PCName = & $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File $scriptPath `
		-Verbose:$false
    Write_LogEntry -Message "Externes Script ausgeführt; PCName ermittelt: $($PCName)" -Level "INFO"
} catch {
    Write_LogEntry -Message "Fehler beim Ausführen des externen Scripts $($scriptPath): $($_)" -Level "ERROR"
    Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
    Pause
    Exit
}

$shortcutSourcePath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HASS.Agent\HASS.Agent.lnk"
$shortcutDestinationPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HASS.Agent.lnk"
Write_LogEntry -Message "Shortcut Source: $($shortcutSourcePath); Destination: $($shortcutDestinationPath)" -Level "DEBUG"
if (Test-Path $shortcutSourcePath) {
	Write_LogEntry -Message "Verschiebe Startmenu Shortcut von $($shortcutSourcePath) nach $($shortcutDestinationPath)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird verschoben." -foregroundcolor "Cyan"
	Move-Item -Path $shortcutSourcePath -Destination $shortcutDestinationPath -Force
    Write_LogEntry -Message "Shortcut verschoben: $($shortcutDestinationPath)" -Level "SUCCESS"
}

$shortcutFolderPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HASS.Agent"
Write_LogEntry -Message "Prüfe Shortcut-Ordner: $($shortcutFolderPath)" -Level "DEBUG"
if (Test-Path $shortcutFolderPath) {
	Write_LogEntry -Message "Entferne Shortcut-Ordner: $($shortcutFolderPath)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
	Remove-Item -Path $shortcutFolderPath -Recurse -Force
    Write_LogEntry -Message "Shortcut-Ordner entfernt: $($shortcutFolderPath)" -Level "SUCCESS"
}

if ($InstallationFlag -eq $true) {
	Write_LogEntry -Message "InstallationFlag = true: Passe Aufgabenplanung an und stelle Konfiguration wieder her" -Level "INFO"
	Write-Host "Aufgabenplannung wird angepasst." -foregroundcolor "Yellow"

	$taskSchedulerScript = "$Serverip\Daten\Customize_Windows\Scripte\HassAgentTaskScheduler.ps1"
    Write_LogEntry -Message "TaskScheduler Script Pfad: $($taskSchedulerScript)" -Level "DEBUG"
	#powershell -ExecutionPolicy Bypass -NoLogo -NoProfile -File $taskSchedulerScript
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File $taskSchedulerScript
    Write_LogEntry -Message "TaskScheduler Script aufgerufen: $($taskSchedulerScript)" -Level "SUCCESS"

	if($PCName -ne "Unknown")
	{
        Write_LogEntry -Message "PCName = $($PCName). Versuche Konfigurationen wiederherzustellen." -Level "INFO"
		Write-Host "	Konfiguration wird zurückgespielt." -foregroundcolor "Yellow"
		
		$sourceConfig = "$Serverip\Daten\Projekte\Smart_Home\HASS_Agent\$PCName\HASS_config"
		#$destinationConfig = "$env:USERPROFILE\AppData\Roaming\LAB02 Research\HASS.Agent\config"
		$destinationConfig = "$env:USERPROFILE\AppData\Local\HASS.Agent\Client\config"
        Write_LogEntry -Message "Quelle Config: $($sourceConfig); Ziel Config: $($destinationConfig)" -Level "DEBUG"
		#Copy-Item -Path $sourceConfig -Destination $destinationConfig -Recurse -Force
		
		if (Test-Path $sourceConfig) {
			Write_LogEntry -Message "Konfig-Backup gefunden: $($sourceConfig). Starte Kopie nach $($destinationConfig)" -Level "INFO"
			Write-Host "	Backup wird wiederhergestellt." -foregroundcolor "Cyan"
			if (!(Test-Path $destinationConfig)) {
				New-Item -ItemType Directory -Path $destinationConfig -Force
                Write_LogEntry -Message "Erstelltes Zielverzeichnis: $($destinationConfig)" -Level "DEBUG"
			}
			Get-ChildItem $sourceConfig | Copy-Item -Destination $destinationConfig -Recurse -Force
            Write_LogEntry -Message "Konfiguration kopiert von $($sourceConfig) nach $($destinationConfig)" -Level "SUCCESS"
		} else {
            Write_LogEntry -Message "Kein Config-Backup gefunden unter: $($sourceConfig)" -Level "WARNING"
		}

		$sourceSatelliteConfig = "$Serverip\Daten\Projekte\Smart_Home\HASS_Agent\$PCName\Satellite_config"
		#$destinationSatelliteConfig = "C:\Program Files (x86)\LAB02 Research\HASS.Agent Satellite Service\config"
		$destinationSatelliteConfig = "C:\Program Files\HASS.Agent\Service\config"
        Write_LogEntry -Message "Quelle Satellite Config: $($sourceSatelliteConfig); Ziel: $($destinationSatelliteConfig)" -Level "DEBUG"
		#Copy-Item -Path $sourceSatelliteConfig -Destination $destinationSatelliteConfig -Recurse -Force
		
		if (Test-Path $sourceSatelliteConfig) {
			Write_LogEntry -Message "Satellite Backup gefunden: $($sourceSatelliteConfig). Starte Kopie nach $($destinationSatelliteConfig)" -Level "INFO"
			Write-Host "	Backup wird wiederhergestellt." -foregroundcolor "Cyan"
			if (!(Test-Path $destinationSatelliteConfig)) {
				New-Item -ItemType Directory -Path $destinationSatelliteConfig -Force
                Write_LogEntry -Message "Erstelltes Zielverzeichnis für Satellite: $($destinationSatelliteConfig)" -Level "DEBUG"
			}
			Get-ChildItem $sourceSatelliteConfig | Copy-Item -Destination $destinationSatelliteConfig -Recurse -Force
            Write_LogEntry -Message "Satellite Konfiguration kopiert von $($sourceSatelliteConfig) nach $($destinationSatelliteConfig)" -Level "SUCCESS"
		} else {
            Write_LogEntry -Message "Kein Satellite-Backup gefunden unter: $($sourceSatelliteConfig)" -Level "WARNING"
		}
		
		# Extract the file version from the filename
		$fileName = [System.IO.Path]::GetFileNameWithoutExtension($agentFile)
		$fileVersionPattern = 'HASS\.Agent\.Installer_(\d+)'
		$localVersion = $fileName -replace $fileVersionPattern, '$1'
        Write_LogEntry -Message "Extrahierte lokale Version aus Dateiname: $($fileName) => $($localVersion)" -Level "DEBUG"

		#Fix für Mqtt Retain
		$shortcutSourcePath = "$Serverip\Daten\Projekte\Smart_Home\HASS_Agent\HASS.Agent.Shared.dll"
		$shortcutDestinationPath = "$env:USERPROFILE\AppData\Roaming\LAB02 Research\HASS.Agent"
        Write_LogEntry -Message "Prüfe Mqtt Retain Fix Quelle: $($shortcutSourcePath); Ziel: $($shortcutDestinationPath)" -Level "DEBUG"
		#2022140 version 22.14.0. Testen, da DLL nur für diese Version
		if ((Test-Path $shortcutSourcePath) -and ($localVersion -eq "2022140")) {
			Write_LogEntry -Message "Mqtt Retain Fix anwendbar (Version $($localVersion)). Kopiere DLL." -Level "INFO"
			Write-Host "	KrX Mqtt Retain Fix." -foregroundcolor "Cyan"
			Copy-Item -Path $shortcutSourcePath -Destination $shortcutDestinationPath -Force
            Write_LogEntry -Message "Mqtt Retain DLL kopiert nach: $($shortcutDestinationPath)" -Level "SUCCESS"
		} else {
            Write_LogEntry -Message "Mqtt Retain Fix nicht angewendet. Bedingung nicht erfüllt oder Datei fehlt." -Level "DEBUG"
		}

		Write-Host "		Konfigurationdateien werden angepasst.." -foregroundcolor "Yellow"
        Write_LogEntry -Message "Beginn Ersetzen in JSON-Dateien: Ersetze 'KrX-HP-Desktop' durch $($PCName)" -Level "INFO"
		# Replace "KRX-HP-Desktop" with "Elitebook-G3" in .json files
		$jsonFiles = Get-ChildItem -Path $destinationConfig -Filter "*.json" -Recurse
		foreach ($jsonFile in $jsonFiles) {
            Write_LogEntry -Message "Verarbeite JSON-Datei: $($jsonFile.FullName)" -Level "DEBUG"
			$content = Get-Content -Path $jsonFile.FullName -Raw
			$updatedContent = $content -replace "KrX-HP-Desktop", $PCName
			Set-Content -Path $jsonFile.FullName -Value $updatedContent
            Write_LogEntry -Message "Aktualisierte JSON-Datei: $($jsonFile.FullName)" -Level "SUCCESS"
		}

		$jsonFiles = Get-ChildItem -Path $destinationSatelliteConfig -Filter "*.json" -Recurse
		foreach ($jsonFile in $jsonFiles) {
            Write_LogEntry -Message "Verarbeite Satellite JSON-Datei: $($jsonFile.FullName)" -Level "DEBUG"
			$content = Get-Content -Path $jsonFile.FullName -Raw
			$updatedContent = $content -replace "KrX-HP-Desktop", $PCName
			Set-Content -Path $jsonFile.FullName -Value $updatedContent
            Write_LogEntry -Message "Aktualisierte Satellite JSON-Datei: $($jsonFile.FullName)" -Level "SUCCESS"
		}
	}
}

# Then start it if it's not running
Write_LogEntry -Message "Starte/prüfe HASS.Agent Dienst/Prozess am Ende des Scripts." -Level "INFO"
Start-HassAgent

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
