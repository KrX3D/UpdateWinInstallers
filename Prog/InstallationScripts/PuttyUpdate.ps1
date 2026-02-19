param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Putty"
$ScriptType  = "Update"

# === Logger-Header: automatisch eingefügt ===
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Logger\Logger.psm1"

if (Test-Path $modulePath) {
    Import-Module -Name $modulePath -Force -ErrorAction Stop

    if (-not (Get-Variable -Name logRoot -Scope Script -ErrorAction SilentlyContinue)) {
        $logRoot = Join-Path -Path $PSScriptRoot -ChildPath "Log"
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
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
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
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Berechneter Konfigurationspfad: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei geladen: $($configPath)" -Level "INFO"
} else {
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    exit
}

$puttyPath = "$InstallationFolder\putty.exe"
$destinationPath = [Environment]::GetFolderPath("Desktop")
Write_LogEntry -Message "Putty Pfad: $($puttyPath); Desktop Zielpfad: $($destinationPath)" -Level "DEBUG"

$installDir = "C:\Program Files (x86)\PuTTY"
$puttyInstallPath = Join-Path -Path $installDir -ChildPath "putty.exe"
$desktopShortcutPath = Join-Path -Path $destinationPath -ChildPath "PuTTY.lnk"

# Function to parse the version number from the release string
function ParseVersion {
    param (
        [string]$VersionString
    )

    $versionPattern = '(\d+\.\d+)'
    $match = [regex]::Match($VersionString, $versionPattern)
    
    if ($match.Success) {
        return $match.Groups[1].Value
    }
    
    return $null
}

# Function to check if a newer version is available
function CheckPuTTYVersion {
    param (
        [string]$InstalledVersion
    )

    $url = "https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html"
    Write_LogEntry -Message "Prüfe Online-Version für PuTTY unter: $($url)" -Level "DEBUG"
    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing
        Write_LogEntry -Message "Online-Seite abgerufen: $($url)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen der PuTTY-Seite $($url): $($_)" -Level "ERROR"
        return
    }

	$versionPattern = 'latest release \((.*?)\)'
	try {
		$latestVersion = [regex]::Match($response.Content, $versionPattern).Groups[1].Value
		Write_LogEntry -Message "Gefundene Online-Version: $($latestVersion)" -Level "DEBUG"
	} catch {
		Write_LogEntry -Message "Fehler beim Parsen der Online-Version aus der Webseite: $($_)" -Level "ERROR"
		return
	}

	Write-Host ""
	Write-Host "Lokale Version: $InstalledVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
	Write-Host ""
	Write_LogEntry -Message "Lokale Version: $($InstalledVersion); Online Version: $($latestVersion)" -Level "INFO"
	
    if ($latestVersion -gt $InstalledVersion) {
		$downloadLinkPattern = '<span class="downloadname">64-bit x86:</span>\s*<span class="downloadfile"><a href="(.*?putty\.exe)">'
		$downloadLink = [regex]::Match($response.Content, $downloadLinkPattern).Groups[1].Value
		Write_LogEntry -Message "Downloadlink extrahiert: $($downloadLink)" -Level "DEBUG"
		
        $downloadPath = (Join-Path -Path $env:TEMP -ChildPath "putty.exe")
        Write_LogEntry -Message "Temporärer Downloadpfad: $($downloadPath)" -Level "DEBUG"
        $webClient = New-Object System.Net.WebClient
        try {
            [void](Invoke-DownloadFile -Url $downloadLink -OutFile $downloadPath)
            Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Herunterladen von $($downloadLink): $($_)" -Level "ERROR"
        } finally {
            $webClient.Dispose()
        }
	 
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
			try {
				if (Test-Path $puttyPath) {
					Remove-Item -Path $puttyPath -Force
					Write_LogEntry -Message "Alte putty.exe auf Quelle entfernt: $($puttyPath)" -Level "DEBUG"
				}
			} catch {
				Write_LogEntry -Message "Fehler beim Entfernen der alten putty.exe auf Quelle $($puttyPath): $($_)" -Level "WARNING"
			}
			
			try {
				Move-Item -Path $downloadPath -Destination $puttyPath -Force
				Write_LogEntry -Message "Neue putty.exe verschoben nach Quelle: $($puttyPath)" -Level "SUCCESS"
			} catch {
				Write_LogEntry -Message "Fehler beim Verschieben der neuen putty.exe nach Quelle $($puttyPath): $($_)" -Level "ERROR"
			}

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
		} else {
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
            Write_LogEntry -Message "Download fehlgeschlagen; temporäre Datei nicht gefunden: $($downloadPath)" -Level "ERROR"
		}		
    } else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Online Update verfügbar für $($ProgramName). Online: $($latestVersion); Lokal: $($InstalledVersion)" -Level "INFO"
    }
}

# Check PuTTY version
try {
    if (Test-Path $puttyInstallPath) {
        $puttyVersionInfo = Get-Item $puttyInstallPath | Select-Object -ExpandProperty VersionInfo
        Write_LogEntry -Message "VersionInfo für lokale putty.exe ermittelt: $($puttyInstallPath)" -Level "DEBUG"
        $installedVersionString = $puttyVersionInfo.ProductVersion
        $installedVersion = ParseVersion -VersionString $installedVersionString
        Write_LogEntry -Message "Installierte VersionString: $($installedVersionString); Parsed: $($installedVersion)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "putty.exe nicht gefunden oder VersionInfo nicht ermittelbar: $($puttyInstallPath)" -Level "WARNING"
        $installedVersion = $null
    }
} catch {
    Write_LogEntry -Message "putty.exe nicht gefunden oder VersionInfo nicht ermittelbar: $($puttyInstallPath); Fehler: $($_)" -Level "WARNING"
    $installedVersion = $null
}

if ($installedVersion) {
    CheckPuTTYVersion -InstalledVersion $installedVersion
} else {
    Write_LogEntry -Message "Putty nicht installiert oder Version unbekannt; überspringe Online-Check." -Level "DEBUG"
    #Write-Host "PuTTY is not installed on this system."
}

Write-Host ""

#Check Installed Version / Install if neded
try {
    $FoundFile = Get-ChildItem $puttyPath
    Write_LogEntry -Message "Gefundene lokale Putty Datei: $($FoundFile.FullName)" -Level "DEBUG"
    $versionInfo = (Get-Item $FoundFile).VersionInfo
    $localVersion = $versionInfo.ProductMajorPart.ToString() + "." + $versionInfo.ProductMinorPart.ToString() + "." + $versionInfo.ProductBuildPart.ToString() + "." + $versionInfo.ProductPrivatePart.ToString()
    Write_LogEntry -Message "Lokale Installationsdatei Version: $($localVersion)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Putty-Datei/Version: $($_)" -Level "WARNING"
    $localVersion = $null
}

$destinationFilePath = Join-Path $env:USERPROFILE "Desktop\putty.exe"
if (Test-Path $destinationFilePath) {
	try {
		$FoundFile = Get-ChildItem $destinationFilePath
		$versionInfo = (Get-Item $FoundFile).VersionInfo
		$installedVersion = $versionInfo.ProductMajorPart.ToString() + "." + $versionInfo.ProductMinorPart.ToString() + "." + $versionInfo.ProductBuildPart.ToString() + "." + $versionInfo.ProductPrivatePart.ToString()
		Write_LogEntry -Message "Gefundene Putty auf Desktop: $($destinationFilePath); Version: $($installedVersion)" -Level "DEBUG"
	} catch {
		Write_LogEntry -Message "Fehler beim Ermitteln der Desktop-Putty-Version: $($_)" -Level "WARNING"
		$installedVersion = $null
	}
} else {
    Write_LogEntry -Message "Putty nicht auf Desktop gefunden: $($destinationFilePath)" -Level "DEBUG"
}

function Normalize-VersionString {
    param([string]$vs)
    if (-not $vs) { return [version]"0.0.0.0" }

    # Extrahiere nur Ziffern-Teile und wandle in ints um (sichere Handhabung, falls zusätzliche Texte vorhanden sind)
    $parts = ($vs -split '\.') | ForEach-Object {
        if ($_ -match '\d+') { [int]($Matches[0]) } else { 0 }
    }

    while ($parts.Count -lt 4) { $parts += 0 }   # auf 4 Teile auffüllen
    # Wenn es mehr als 4 Teile gäbe, werden nur die ersten 4 verwendet
    return [version]::new($parts[0], $parts[1], $parts[2], $parts[3])
}

if ($null -ne $installedVersion) {
    # Normalisiere beide Versionen auf Version-Objekte mit 4 Teilen
    $instVerObj  = Normalize-VersionString $installedVersion
    $localVerObj = Normalize-VersionString $localVersion

    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $($instVerObj.ToString())" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $($localVerObj.ToString())" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$($ProgramName) ist installiert; Installierte Version: $($instVerObj.ToString()); Installationsdatei Version: $($localVerObj.ToString())" -Level "INFO"
    
    if ($instVerObj -lt $localVerObj) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
        $Install = $true
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (Update erforderlich)" -Level "INFO"
    } elseif ($instVerObj -eq $localVerObj) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
        $Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (Version aktuell)" -Level "DEBUG"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
        $Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (installierte Version neuer)" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
    $Install = $false
    Write_LogEntry -Message "$($ProgramName) nicht installiert (keine Version in Zielpfad gefunden)." -Level "INFO"
}
Write-Host ""
Write_LogEntry -Message "Install-Flag/Entscheidung: InstallationFlag=$($InstallationFlag); Install=$($Install)" -Level "DEBUG"

#Install if needed
if($Install -eq $true -or $InstallationFlag)
{
	Write-Host "Putty wird kopiert"
    Write_LogEntry -Message "Starte Kopiervorgang Putty auf Ziel: $($puttyInstallPath) (Quelle: $($puttyPath))" -Level "INFO"

	if (-not (Test-Path -Path $installDir)) {
		try {
			New-Item -Path $installDir -ItemType Directory -Force | Out-Null
            Write_LogEntry -Message "Installationsverzeichnis erstellt: $($installDir)" -Level "INFO"
		} catch {
			Write_LogEntry -Message "Fehler beim Erstellen des Installationsverzeichnisses $($installDir): $($_)" -Level "ERROR"
		}
	}

	if (Test-Path $puttyPath) {
		try {
			Copy-Item -Path $puttyPath -Destination $puttyInstallPath -Force
			Write_LogEntry -Message "Putty kopiert nach: $($puttyInstallPath)" -Level "SUCCESS"
		} catch {
			Write_LogEntry -Message "Fehler beim Kopieren von $($puttyPath) nach $($puttyInstallPath): $($_)" -Level "ERROR"
		}
	} else {
		Write_LogEntry -Message "Quelle Putty nicht gefunden: $($puttyPath)" -Level "ERROR"
	}
	
	# Create Desktop shortcut only when copying/installing
	try {
		$shell = New-Object -ComObject WScript.Shell
		$shortcut = $shell.CreateShortcut($desktopShortcutPath)
		$shortcut.TargetPath = $puttyInstallPath
		$shortcut.WorkingDirectory = $installDir
		if (Test-Path $puttyInstallPath) {
			$shortcut.IconLocation = $puttyInstallPath
		}
		$shortcut.Save()
		Write_LogEntry -Message "Desktop-Verknüpfung erstellt: $($desktopShortcutPath)" -Level "SUCCESS"
	} catch {
		Write_LogEntry -Message "Fehler beim Erstellen der Desktop-Verknüpfung $($desktopShortcutPath): $($_)" -Level "ERROR"
	}

	Write-Host "Putty wird kofiguriert"
	Write_LogEntry -Message "Beginne Konfiguration von Putty (SSH Host Keys & Registry-Einstellungen)" -Level "INFO"
	Write-Host "	SSH Host Keys werden gesetzt"
		#PC Name wird gesucht
		$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
		Write_LogEntry -Message "Rufe PC-Name Script auf: $($scriptPath)" -Level "DEBUG"

		try {
			#$PCName = & $scriptPath
			#$PCName = & "powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
			$PCName = & $PSHostPath `
				-NoLogo -NoProfile -ExecutionPolicy Bypass `
				-File $scriptPath `
				-Verbose:$false
            Write_LogEntry -Message "PC-Name ermittelt: $($PCName) vom Script: $($scriptPath)" -Level "DEBUG"
		} catch {
			Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
            Write_LogEntry -Message "Fehler beim Aufruf des Scripts $($scriptPath): $($($_.Exception.Message))" -Level "ERROR"
			Exit
		}
		
		# Define the backup file path
		$backupFilePath = "$Serverip\Daten\Windows_Backup\$PCName\PuTTY_SSHHostKeys_Backup.reg"
		Write_LogEntry -Message "Erwarteter Backup-Pfad für SSH Host Keys: $($backupFilePath)" -Level "DEBUG"

		# Check if the backup file exists
		if (Test-Path $backupFilePath) {
            try {
			    # Import the SSH host keys from the .reg file
			    reg import $backupFilePath
			    Write-Host "Registry file $backupFilePath imported successfully."
                Write_LogEntry -Message "Registry Datei importiert: $($backupFilePath)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "Fehler beim Import der Registry-Datei $($backupFilePath): $($_)" -Level "ERROR"
            }
		} else {
			Write-Host "Registry file not found: $backupFilePath"
            Write_LogEntry -Message "Registry Backup Datei nicht gefunden: $($backupFilePath)" -Level "WARNING"
		}
	
	Write-Host "	SSH Host Keys werden gesetzt"
	Write_LogEntry -Message "Setze PuTTY Default Settings in Registry: HKCU:\Software\SimonTatham\PuTTY\Sessions\Default%20Settings" -Level "DEBUG"
		# Define the path to the PuTTY default settings in the registry
		$puttyDefaultSettingsPath = "HKCU:\Software\SimonTatham\PuTTY\Sessions\Default%20Settings"

		# Define the font name you want to set
		$fontName = "Terminal"

		# Define the FontCharSet value you want to set (DWORD value in hexadecimal)
		$fontCharSet = 0x000000FF

		# Check if the registry path exists
		if (-not (Test-Path -Path $puttyDefaultSettingsPath)) {
			# Create the registry key if it does not exist
			New-Item -Path "HKCU:\Software\SimonTatham\PuTTY\Sessions" -Name "Default%20Settings" -Force
            Write_LogEntry -Message "Registry-Pfad erstellt: $($puttyDefaultSettingsPath)" -Level "DEBUG"
		}

		# Set the font in the registry
		try {
			Set-ItemProperty -Path $puttyDefaultSettingsPath -Name "Font" -Value $fontName
			Write_LogEntry -Message "Registry-Eintrag gesetzt: Font = $($fontName) in $($puttyDefaultSettingsPath)" -Level "DEBUG"
		} catch {
			Write_LogEntry -Message "Fehler beim Setzen des Registry-Eintrags Font in $($puttyDefaultSettingsPath): $($_)" -Level "ERROR"
		}

		# Set the FontCharSet in the registry
		try {
			Set-ItemProperty -Path $puttyDefaultSettingsPath -Name "FontCharSet" -Value $fontCharSet -Type DWord
			Write_LogEntry -Message "Registry-Eintrag gesetzt: FontCharSet = $($fontCharSet) in $($puttyDefaultSettingsPath)" -Level "DEBUG"
		} catch {
			Write_LogEntry -Message "Fehler beim Setzen des Registry-Eintrags FontCharSet in $($puttyDefaultSettingsPath): $($_)" -Level "ERROR"
		}
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===
