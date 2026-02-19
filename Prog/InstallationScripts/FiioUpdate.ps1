param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "FiiO"
$ScriptType = "Update"

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

Write_LogEntry -Message "Lade Konfigurationsdatei von: $configPath" -Level "INFO"

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

$skipDownload = $false

# Define the FiiO driver folder path
$FiiODriverFolder = Join-Path -Path $NetworkShareDaten -ChildPath "Treiber\FiiO_Verstarker"
Write_LogEntry -Message "FiiO Treiber-Ordner: $FiiODriverFolder" -Level "INFO"

# Define the URL of the FiiO driver page
$downloadPageUrl = "https://forum.fiio.com/note/showNoteContent.do?id=202105191527366657910&tid=17"
Write_LogEntry -Message "Download-Seite URL: $downloadPageUrl" -Level "INFO"

# Search for local installation file
$InstallationFilePattern = Join-Path -Path $FiiODriverFolder -ChildPath "FiiO_v*.msi"
Write_LogEntry -Message "Suche nach lokaler Installationsdatei: $InstallationFilePattern" -Level "INFO"

$FoundFile = Get-ChildItem -Path $InstallationFilePattern -ErrorAction SilentlyContinue
if ($FoundFile) {
    $InstallationFileName = ($FoundFile | Select-Object -First 1).Name
    $localInstaller = Join-Path -Path $FiiODriverFolder -ChildPath $InstallationFileName
    Write_LogEntry -Message "Lokale Installationsdatei gefunden: $InstallationFileName" -Level "SUCCESS"
    
    # Extract version from filename (e.g., FiiO_v5.50.0_2022-12-02.msi -> 5.50.0)
    if ($InstallationFileName -match 'FiiO_v([\d\.]+)') {
        $localVersion = $matches[1]
        Write_LogEntry -Message "Lokale Version aus Dateiname ermittelt: $localVersion" -Level "INFO"
    } else {
        $localVersion = "0.0.0"
        Write_LogEntry -Message "Fehler beim Auslesen der lokalen Version aus Dateiname" -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei gefunden - Download-Teil wird übersprungen, aber Registry-Check läuft weiter" -Level "ERROR"
    $localVersion = "0.0.0"
    $skipDownload = $true
}

if (!$skipDownload) {
    # Retrieve the download page content
    Write_LogEntry -Message "Lade Download-Seite herunter..." -Level "INFO"
    try {
        $webRequest = Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing
        $pageContent = $webRequest.Content
        Write_LogEntry -Message "Download-Seite erfolgreich abgerufen (Größe: $($pageContent.Length) Zeichen)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen der Download-Seite: $_" -Level "ERROR"
        $skipDownload = $true
    }

    if (!$skipDownload) {
        # Extract version and download link for Win 10/11 version
        Write_LogEntry -Message "Analysiere Seiteninhalt für Download-Links..." -Level "INFO"
        
        # Pattern to find version number (e.g., V5.74.3)
        $versionPattern = 'V([\d\.]+)\s+version\s+download\s+link:.*?<a[^>]*href="([^"]+)"[^>]*>.*?Click here.*?</a>.*?\(For Win 10/11'
        
        $versionMatch = [regex]::Match($pageContent, $versionPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        if ($versionMatch.Success) {
            $onlineVersion = $versionMatch.Groups[1].Value
            $downloadLink = $versionMatch.Groups[2].Value
            
            Write_LogEntry -Message "Online Version gefunden: $onlineVersion" -Level "SUCCESS"
            Write_LogEntry -Message "Download-Link gefunden: $downloadLink" -Level "SUCCESS"
            
            Write-Host ""
            Write-Host "Lokale Version: $localVersion" -ForegroundColor "Cyan"
            Write-Host "Online Version: $onlineVersion" -ForegroundColor "Cyan"
            Write-Host ""
            
            Write_LogEntry -Message "Versionsvergleich - Lokal: $localVersion, Online: $onlineVersion" -Level "INFO"
            
            # Compare versions
            try {
                $localVersionObj = [version]$localVersion
                $onlineVersionObj = [version]$onlineVersion
                
                if ($onlineVersionObj -le $localVersionObj) {
                    Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
                    Write_LogEntry -Message "Keine neuere Version verfügbar - $ProgramName ist aktuell" -Level "INFO"
                } else {
                    Write_LogEntry -Message "Neuere Version verfügbar - starte Download..." -Level "INFO"
                    
                    # Construct download URL if link is relative or use direct link
                    if ($downloadLink -notmatch '^https?://') {
                        # Fallback: construct download URL
                        $downloadUrl = "https://fiio-firmware.fiio.net/DAC/FiiO%20USB%20DAC%20driver%20V$onlineVersion.exe"
                        Write_LogEntry -Message "Konstruierte Download-URL: $downloadUrl" -Level "INFO"
                    } else {
                        $downloadUrl = $downloadLink
                        Write_LogEntry -Message "Verwende extrahierten Download-Link: $downloadUrl" -Level "INFO"
                    }
                    
                    # Download to temp folder first
                    $tempFolder = Join-Path -Path $env:TEMP -ChildPath "FiiO_Update_$onlineVersion"
                    if (!(Test-Path $tempFolder)) {
                        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
                        Write_LogEntry -Message "Temporärer Ordner erstellt: $tempFolder" -Level "INFO"
                    }
                    
                    $tempExePath = Join-Path -Path $tempFolder -ChildPath "FiiO_USB_DAC_driver_V$onlineVersion.exe"
                    Write_LogEntry -Message "Ziel-Pfad für Download: $tempExePath" -Level "INFO"
                    
                    try {
                        Write_LogEntry -Message "Starte Download von $downloadUrl..." -Level "INFO"
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($downloadUrl, $tempExePath)
                        $webClient.Dispose()
                        Write_LogEntry -Message "Download erfolgreich abgeschlossen" -Level "SUCCESS"
                    } catch {
                        Write_LogEntry -Message "Fehler beim Download: $_" -Level "ERROR"
                        $skipDownload = $true
                    }
                    
                    if (!$skipDownload) {
                        # Check if the file was completely downloaded
                        if (Test-Path $tempExePath) {
                            $downloadedFileSize = (Get-Item $tempExePath).Length
                            Write_LogEntry -Message "Download-Datei verifiziert (Größe: $downloadedFileSize Bytes)" -Level "SUCCESS"
                            
                            # Extract the EXE file to get the MSI
                            Write_LogEntry -Message "Extrahiere MSI aus EXE-Datei..." -Level "INFO"
                            $extractFolder = Join-Path -Path $tempFolder -ChildPath "Extracted"
                            
                            try {
                                # Create extraction folder
                                if (!(Test-Path $extractFolder)) {
                                    New-Item -Path $extractFolder -ItemType Directory -Force | Out-Null
                                }
                                
                                # Use 7-Zip or native extraction
                                $sevenZipPath = "C:\Program Files\7-Zip\7z.exe"
                                if (Test-Path $sevenZipPath) {
                                    Write_LogEntry -Message "Verwende 7-Zip für Extraktion" -Level "INFO"
                                    $extractResult = & $sevenZipPath x "$tempExePath" "-o$extractFolder" -y 2>&1
                                    Write_LogEntry -Message "7-Zip Extraktion abgeschlossen" -Level "INFO"
                                } else {
                                    Write_LogEntry -Message "7-Zip nicht gefunden, verwende alternative Methode" -Level "WARNING"
                                    # Alternative: try to run the EXE with silent extraction parameters
                                    [void](Invoke-InstallerFile -FilePath $tempExePath -Arguments "/S", "/D=$extractFolder" -Wait)
                                }
                                
                                # Find the MSI file in the x64 folder
                                $msiPath = Join-Path -Path $extractFolder -ChildPath "x64\*.msi"
                                $extractedMsi = Get-ChildItem -Path $msiPath -ErrorAction SilentlyContinue | Select-Object -First 1
                                
                                if ($extractedMsi) {
                                    Write_LogEntry -Message "MSI-Datei gefunden: $($extractedMsi.Name)" -Level "SUCCESS"
                                    
                                    # Get current date for filename
                                    $currentDate = Get-Date -Format "yyyy-MM-dd"
                                    $newMsiName = "FiiO_v${onlineVersion}_${currentDate}.msi"
                                    $newMsiPath = Join-Path -Path $FiiODriverFolder -ChildPath $newMsiName
                                    
                                    # Copy MSI to driver folder on NAS
                                    Write_LogEntry -Message "Kopiere MSI zur NAS: $newMsiPath" -Level "INFO"
                                    Copy-Item -Path $extractedMsi.FullName -Destination $newMsiPath -Force
                                    Write_LogEntry -Message "MSI-Datei erfolgreich kopiert nach: $newMsiPath" -Level "SUCCESS"
                                    
                                    # Verify the copy was successful
                                    if (Test-Path $newMsiPath) {
                                        # Remove old installer only after successful copy
                                        if (Test-Path $localInstaller) {
                                            Write_LogEntry -Message "Entferne alte Installationsdatei: $localInstaller" -Level "INFO"
                                            Remove-Item $localInstaller -Force
                                            Write_LogEntry -Message "Alte Installationsdatei erfolgreich entfernt" -Level "SUCCESS"
                                        }
                                        
                                        Write-Host "$ProgramName wurde aktualisiert." -ForegroundColor "Green"
                                        Write_LogEntry -Message "$ProgramName wurde erfolgreich aktualisiert auf Version $onlineVersion" -Level "SUCCESS"
                                    } else {
                                        Write-Host "Fehler beim Kopieren der MSI-Datei zur NAS." -ForegroundColor "Red"
                                        Write_LogEntry -Message "MSI-Datei konnte nicht zur NAS kopiert werden" -Level "ERROR"
                                    }
                                } else {
                                    Write-Host "MSI-Datei konnte nicht in extrahiertem Ordner gefunden werden." -ForegroundColor "Red"
                                    Write_LogEntry -Message "MSI-Datei nicht im x64-Ordner gefunden: $msiPath" -Level "ERROR"
                                    
                                    # Log what was found in extraction folder for debugging
                                    if (Test-Path $extractFolder) {
                                        $foundFiles = Get-ChildItem -Path $extractFolder -Recurse | Select-Object -ExpandProperty FullName
                                        Write_LogEntry -Message "Gefundene Dateien im Extraktionsordner: $($foundFiles -join ', ')" -Level "DEBUG"
                                    }
                                }
                            } catch {
                                Write_LogEntry -Message "Fehler beim Extrahieren: $_" -Level "ERROR"
                                Write-Host "Fehler beim Extrahieren der MSI-Datei." -ForegroundColor "Red"
                            } finally {
                                # Clean up temporary files
                                Write_LogEntry -Message "Bereinige temporäre Dateien..." -Level "INFO"
                                if (Test-Path $tempFolder) {
                                    Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
                                    Write_LogEntry -Message "Temporäre Dateien entfernt: $tempFolder" -Level "SUCCESS"
                                }
                            }
                        } else {
                            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "Red"
                            Write_LogEntry -Message "Download fehlgeschlagen - Datei nicht gefunden: $tempExePath" -Level "ERROR"
                        }
                    }
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Versionsvergleich: $_" -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Keine passende Version und Download-Link auf der Webseite gefunden" -Level "WARNING"
            Write-Host "Keine passende Version auf der Webseite gefunden." -ForegroundColor "Yellow"
        }
    }
}
Write-Host ""

# Check Installed Version / Install if needed
Write_LogEntry -Message "Prüfe installierte Version von $ProgramName..." -Level "INFO"

# Re-scan for local installer in case it was just updated
$FoundFile = Get-ChildItem -Path $InstallationFilePattern -ErrorAction SilentlyContinue
if ($FoundFile) {
    $InstallationFileName = ($FoundFile | Select-Object -First 1).Name
    $localInstaller = Join-Path -Path $FiiODriverFolder -ChildPath $InstallationFileName
    
    if ($InstallationFileName -match 'FiiO_v([\d\.]+)') {
        $localVersion = $matches[1]
    } else {
        $localVersion = "0.0.0"
    }
    Write_LogEntry -Message "Aktuelle lokale Installationsdatei: $InstallationFileName (Version: $localVersion)" -Level "INFO"
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei für Installationscheck gefunden" -Level "WARNING"
    $localVersion = "0.0.0"
}

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

Write_LogEntry -Message "Durchsuche Registry-Pfade nach installierten Programmen..." -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Prüfe Registry-Pfad: $RegPath" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "*FiiO*" -or $_.DisplayName -like "*USB DAC*" }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht gefunden: $RegPath" -Level "DEBUG"
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName Treiber ist installiert." -ForegroundColor "Green"
    Write-Host "	Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -ForegroundColor "Cyan"
    
    Write_LogEntry -Message "$ProgramName Treiber ist installiert - Version: $installedVersion" -Level "SUCCESS"
    Write_LogEntry -Message "Vergleiche installierte Version ($installedVersion) mit lokaler Datei ($localVersion)" -Level "INFO"
    
    # Compare versions if both are valid
    if ($localVersion -and ($localVersion -match '^\d+(\.\d+)*$') -and $installedVersion -and ($installedVersion -match '^\d+(\.\d+)*$')) {
        try {
            if ([version]$installedVersion -lt [version]$localVersion) {
                Write-Host "        Veralteter $ProgramName Treiber ist installiert. Update wird gestartet." -ForegroundColor "Magenta"
                Write_LogEntry -Message "Installierte Version ist veraltet - Update erforderlich" -Level "WARNING"
                $Install = $true
            } elseif ([version]$installedVersion -eq [version]$localVersion) {
                Write-Host "        Installierte Version ist aktuell." -ForegroundColor "DarkGray"
                Write_LogEntry -Message "Installierte Version ist bereits aktuell" -Level "INFO"
                $Install = $false
            } else {
                Write_LogEntry -Message "Installierte Version ($installedVersion) ist höher als lokale Version ($localVersion)" -Level "INFO"
                $Install = $false
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Versionsvergleich: $_" -Level "ERROR"
            $Install = $false
        }
    } else {
        Write_LogEntry -Message "Versionsvergleich nicht möglich - ungültige Versionsnummern" -Level "WARNING"
        $Install = $false
    }
} else {
    Write_LogEntry -Message "$ProgramName Treiber ist nicht auf diesem System installiert" -Level "INFO"
    $Install = $false
}
Write-Host ""

# Install if needed
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt - starte Installation..." -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\FiiOInstallation.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Installation-Script aufgerufen mit InstallationFlag" -Level "INFO"
}
elseif ($Install -eq $true) {
    Write_LogEntry -Message "Update erforderlich - starte Installation..." -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\FiiOInstallation.ps1"
    Write_LogEntry -Message "Installation-Script für Update aufgerufen" -Level "INFO"
} else {
    Write_LogEntry -Message "Keine Installation oder Update erforderlich" -Level "INFO"
}
Write-Host ""

# Finalize logging
Finalize_LogSession -FinalizeMessage "FiiO Treiber Update-Script erfolgreich abgeschlossen"
