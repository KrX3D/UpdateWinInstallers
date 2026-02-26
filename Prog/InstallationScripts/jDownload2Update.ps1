param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "jDownloader 2"
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

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei geladen: $($configPath)" -Level "DEBUG"
} else {
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    exit
}

$InstallationFile = "$InstallationFolder\JDownloader2*.exe"

# Specify the path where you want to save the downloaded installer
$FoundFile = Get-ChildItem -Path $InstallationFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
if ($null -eq $FoundFile) {
    Write_LogEntry -Message "Keine Installationsdatei gefunden mit Muster: $($InstallationFile)" -Level "WARNING"
    # leave variables blank to avoid exceptions later; the script will continue (install logic may still run if flag set)
    $InstallationFileName = $null
    $localInstaller = $null
    $localLastWriteTime = $null
} else {
    $InstallationFileName = $FoundFile.Name
    $localInstaller = "$InstallationFolder\$InstallationFileName"

    # Get the local JDownloader installer's last write time (safely)
    try {
        $localLastWriteTime = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).LastWriteTime
    } catch {
        $localLastWriteTime = $null
        Write_LogEntry -Message "Konnte LastWriteTime der lokalen Installationsdatei nicht bestimmen: $($_.Exception.Message)" -Level "WARNING"
    }
}

Write_LogEntry -Message "Lokale Installationsdatei: $($localInstaller) (LastWriteTime: $($localLastWriteTime))" -Level "DEBUG"

# Define the URL of the JDownloader build website
$buildWebsiteUrl = "https://svn.jdownloader.org/build.php"

# Define the URL of the JDownloader website
$websiteUrl = "https://jdownloader.org/jdownloader2"

# Download the content of the JDownloader build website
try {
    $webContent = Invoke-WebRequest $buildWebsiteUrl -UseBasicParsing -ErrorAction Stop | Select-Object -ExpandProperty Content
    Write_LogEntry -Message "Build-Website abgerufen: $($buildWebsiteUrl)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abruf der Build-Website: $($_.Exception.Message)" -Level "ERROR"
    $webContent = $null
}

$onlineLastWriteTime = $null

if ($webContent) {
    # Extract the date and time from the website content
    $pattern = '(\w{3} \w{3} \d{2} \d{2}:\d{2}:\d{2} (CEST|CET) \d{4})'
    $dateMatch = [regex]::Match($webContent, $pattern)

    if ($dateMatch.Success) {
        $onlineDateTime = $dateMatch.Groups[1].Value
        Write_LogEntry -Message "Rohes Datum von Website: $($onlineDateTime)" -Level "DEBUG"

        # --- Robustere Datums-Parsing & Fallbacks (ParseExact in try/catch statt TryParseExact) ---
        $parseSucceeded = $false
        $formats = @(
            "ddd MMM dd HH:mm:ss 'CEST' yyyy",
            "ddd MMM dd HH:mm:ss 'CET' yyyy",
            "ddd MMM dd HH:mm:ss yyyy"
        )

        foreach ($fmt in $formats) {
            try {
                $onlineLastWriteTime = [DateTime]::ParseExact($onlineDateTime, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
                $parseSucceeded = $true
                break
            } catch {
                # ignore and try next format
            }
        }

        if (-not $parseSucceeded) {
            # Entferne TZ-Token (falls unsichtbare Zeichen o.ä.) und nochmal versuchen
            $clean = $onlineDateTime -replace '\s+(CEST|CET)\s+', ' '
            Write_LogEntry -Message "Versuche geparstes Datum ohne TZ: $($clean)" -Level "DEBUG"
            try {
                $onlineLastWriteTime = [DateTime]::ParseExact($clean, "ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                $parseSucceeded = $true
            } catch {
                # failed
            }
        }

        if ($parseSucceeded) {
            Write-Host ""
            Write-Host "Lokale Version: $localLastWriteTime" -foregroundcolor "Cyan"
            Write-Host "Online Version: $onlineLastWriteTime" -foregroundcolor "Cyan"
            Write_LogEntry -Message "Lokale Version: $($localLastWriteTime); Online Version: $($onlineLastWriteTime)" -Level "INFO"
        } else {
            Write-Host ""
            Write-Host "WARN: Online-Datum konnte nicht geparst werden: $onlineDateTime" -ForegroundColor "Yellow"
            Write_LogEntry -Message "Unable to parse the online date and time: $($onlineDateTime)" -Level "WARNING"
            # important: do NOT return; continue script. $onlineLastWriteTime stays $null
        }
    } else {
        Write_LogEntry -Message "Datum/Zeit konnte nicht von der JDownloader Build-Website extrahiert werden." -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Kein Inhalt von Build-Website erhalten; überspringe Datumsextraktion." -Level "WARNING"
}

Write-Host ""
# Compare local and online last write times (only if $onlineLastWriteTime is available)
$updateAvailable = $false
if ($localLastWriteTime -and $onlineLastWriteTime) {
    if ($localLastWriteTime -lt $onlineLastWriteTime) {
        Write_LogEntry -Message "Online-Version neuer als lokale Version. Update verfügbar." -Level "INFO"
        $updateAvailable = $true
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Online Update verfügbar. $($ProgramName) ist aktuell." -Level "INFO"
    }
} else {
    Write_LogEntry -Message "Online-Datum unbekannt oder lokale Datei nicht vorhanden; überspringe Online-Vergleich." -Level "WARNING"
    # $updateAvailable bleibt $false; we still continue script
}

# If an update is available, attempt to download; if parsing failed we skip automatic download path
if ($updateAvailable) {
    # Download the content of the JDownloader website to find the MEGA link
    try {
        $webContent = Invoke-WebRequest $websiteUrl -UseBasicParsing -ErrorAction Stop | Select-Object -ExpandProperty Content
        Write_LogEntry -Message "JDownloader Website abgerufen: $($websiteUrl)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Abruf der JDownloader-Website: $($_.Exception.Message)" -Level "ERROR"
        $webContent = $null
    }

    if ($webContent) {
        # Extract the download link from the website content
        $pattern = '<td class="col2"> <a id="windows0" href="([^"]+)" class="urlextern" target="_blank"'
        $linkMatch = [regex]::Match($webContent, $pattern)

        if ($linkMatch.Success) {
            $downloadLink = $linkMatch.Groups[1].Value
            Write_LogEntry -Message "Downloadlink gefunden: $($downloadLink)" -Level "DEBUG"

            # Define source and destination paths
            $sourcePath = $InstallationFolder + "\InstallationScripts\MEGAcmd"
            $destinationPath = Join-Path $env:LOCALAPPDATA "MEGAcmd"

            # Copy the MEGAcmd folder to the destination path
            Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse -Force -ErrorAction SilentlyContinue
            Write_LogEntry -Message "MEGAcmd kopiert von $($sourcePath) nach $($destinationPath)" -Level "DEBUG"

            # Set the MEGAcmd executable path
            $megaCmdPath = Join-Path $destinationPath "MEGAclient.exe"

            # Define the download command
            $downloadCommand = "$megaCmdPath get $downloadLink $env:TEMP"

            # Download the JDownloader installer
            Write_LogEntry -Message "Starte Download via MEGAcmd: $($downloadCommand)" -Level "INFO"
            try {
                $cmdProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $downloadCommand -NoNewWindow -PassThru -ErrorAction Stop
                $cmdProcess.WaitForExit()
                Write_LogEntry -Message "MEGAcmd Download Prozess beendet (ExitCode: $($cmdProcess.ExitCode))" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "MEGAcmd Download Prozess fehlgeschlagen: $($_.Exception.Message)" -Level "ERROR"
            }

            Write-Host ""

            # Check if MEGAcmdServer process is running and kill it
            $megaCmdServerProcess = Get-Process -Name "MEGAcmdServer" -ErrorAction SilentlyContinue
            if ($megaCmdServerProcess) {
                $megaCmdServerProcess | Stop-Process -Force -ErrorAction SilentlyContinue
                Write_LogEntry -Message "MEGAcmdServer Prozess beendet" -Level "DEBUG"
            }

            Start-Sleep -Seconds 3

            # Remove the MEGAcmd folder from the destination path
            Remove-Item -Path $destinationPath -Recurse -Force -ErrorAction SilentlyContinue
            Write_LogEntry -Message "MEGAcmd Ordner entfernt: $($destinationPath)" -Level "DEBUG"

            # Copy downloaded file to NAS
            $tempFileObj = Get-ChildItem -Path "$env:TEMP\JDownloader2*.exe" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($tempFileObj) {
                $tempFilePath = $tempFileObj.FullName
                Write_LogEntry -Message "Temporäre Datei gefunden: $($tempFilePath)" -Level "DEBUG"

                # If onlineLastWriteTime is set, use it as LastWriteTime for the downloaded file
                if ($onlineLastWriteTime) {
                    try {
                        Set-ItemProperty -Path $tempFilePath -Name LastWriteTime -Value $onlineLastWriteTime -ErrorAction Stop
                        Write_LogEntry -Message "LastWriteTime der temporären Datei gesetzt: $($onlineLastWriteTime)" -Level "DEBUG"
                    } catch {
                        Write_LogEntry -Message "Konnte LastWriteTime der temporären Datei nicht setzen: $($_.Exception.Message)" -Level "WARNING"
                    }
                }

                # Remove the old jDownloader exe file (if present)
                if ($localInstaller) {
                    try {
                        Remove-Item -Path $localInstaller -Recurse -Force -ErrorAction SilentlyContinue
                        Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localInstaller)" -Level "DEBUG"
                    } catch {
                        Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei: $($_.Exception.Message)" -Level "WARNING"
                    }
                }

                # Move the file to the download path
                try {
                    Move-Item -Path $tempFilePath -Destination $InstallationFolder -Force -ErrorAction Stop
                    Write_LogEntry -Message "Neue Installationsdatei verschoben nach: $($InstallationFolder)" -Level "SUCCESS"
                    Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                    Write_LogEntry -Message "$($ProgramName) wurde aktualisiert.." -Level "SUCCESS"
                } catch {
                    Write_LogEntry -Message "Fehler beim Verschieben der temporären Datei: $($_.Exception.Message)" -Level "ERROR"
                }
            } else {
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "red"
                Write_LogEntry -Message "Download ist fehlgeschlagen. Temporäre Datei nicht gefunden." -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Downloadlink von der JDownloader-Website konnte nicht extrahiert werden." -Level "WARNING"
        }
    } else {
        Write_LogEntry -Message "Website für Downloadlink konnte nicht abgerufen werden; überspringe Download." -Level "WARNING"
    }
}

Write_LogEntry -Message "Update-Check abgeschlossen." -Level "DEBUG"

#Check Installed Version / Install if needed
# Re-evaluate local installer file (in case it was updated above)
$FoundFile = Get-ChildItem -Path $InstallationFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
if ($FoundFile) {
    $InstallationFileName = $FoundFile.Name
    $localInstaller = "$InstallationFolder\$InstallationFileName"
    try {
        $localLastWriteTime = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).LastWriteTime
    } catch {
        $localLastWriteTime = $null
        Write_LogEntry -Message "Konnte LastWriteTime nach Update nicht bestimmen: $($_.Exception.Message)" -Level "WARNING"
    }
    Write_LogEntry -Message "Erneut lokale Installationsdatei bestimmt: $($localInstaller) (LastWriteTime: $($localLastWriteTime))" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine Installationsdatei gefunden beim erneuten Check." -Level "DEBUG"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Prüfe Registry-Pfad: $RegPath" -Level "DEBUG"
        try {
            Get-ChildItem $RegPath -ErrorAction Stop | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
        } catch {
            Write_LogEntry -Message "Fehler beim Zugriff auf Registry-Pfad $RegPath : $($_.Exception.Message)" -Level "WARNING"
        }
    } else {
        Write_LogEntry -Message "Registry-Pfad existiert nicht: $RegPath" -Level "DEBUG"
    }
}

$Install = $false

if ($null -ne $Path) {
    $installedVersion = $Path.TimestampVersion | Select-Object -First 1

    # Try to parse installedVersion in a robust way if present
    if ($installedVersion) {
        Write_LogEntry -Message "TimestampVersion in Registry gefunden: $($installedVersion)" -Level "DEBUG"
        
        $currentTimeZone = if ((Get-Date).IsDaylightSavingTime()) { "CEST" } else { "CET" }
        try {
            $installedVersion = [DateTime]::ParseExact($installedVersion, "ddd MMM dd HH:mm:ss '$currentTimeZone' yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
            Write_LogEntry -Message "TimestampVersion erfolgreich geparst: $($installedVersion)" -Level "DEBUG"
        } catch {
            # Fallback: try to parse without TZ token
            try {
                $installedVersion = [DateTime]::ParseExact(($installedVersion -replace '\s+(CEST|CET)\s+', ' '), "ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                Write_LogEntry -Message "TimestampVersion mit Fallback-Methode geparst: $($installedVersion)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Konnte installierte Version aus Registry nicht parsen: $($installedVersion)" -Level "WARNING"
                $installedVersion = $null
            }
        }
    } else {
        # TimestampVersion nicht in Registry gefunden - versuche build.json
        Write_LogEntry -Message "TimestampVersion nicht in Registry gefunden. Versuche build.json zu lesen..." -Level "WARNING"
        
        $filePath = "C:\Program Files\JDownloader\build.json"
        Write_LogEntry -Message "Lese Build-Info aus: $($filePath)" -Level "DEBUG"
        
        if (Test-Path -Path $filePath) {
            try {
                $jsonContent = Get-Content -Raw -Path $filePath | ConvertFrom-Json
                # Extract the value of 'buildDate'
                $buildTimestamp = $jsonContent.buildDate
                
                if ($buildTimestamp) {
                    Write_LogEntry -Message "buildDate aus build.json gelesen: $($buildTimestamp)" -Level "INFO"
                    
                    # Parse the buildDate to DateTime
                    $currentTimeZone = if ((Get-Date).IsDaylightSavingTime()) { "CEST" } else { "CET" }
                    try {
                        $installedVersion = [DateTime]::ParseExact($buildTimestamp, "ddd MMM dd HH:mm:ss '$currentTimeZone' yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                        Write_LogEntry -Message "buildDate erfolgreich geparst: $($installedVersion)" -Level "DEBUG"
                    } catch {
                        try {
                            $installedVersion = [DateTime]::ParseExact(($buildTimestamp -replace '\s+(CEST|CET)\s+', ' '), "ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                            Write_LogEntry -Message "buildDate mit Fallback-Methode geparst: $($installedVersion)" -Level "DEBUG"
                        } catch {
                            Write_LogEntry -Message "Konnte buildDate nicht parsen: $($buildTimestamp)" -Level "WARNING"
                            $installedVersion = $null
                        }
                    }
                } else {
                    Write_LogEntry -Message "buildDate nicht in $($filePath) gefunden. Setze installedVersion auf null." -Level "WARNING"
                    $installedVersion = $null
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Lesen/Parsen von $($filePath): $($_.Exception.Message)" -Level "ERROR"
                $installedVersion = $null
            }
        } else {
            Write_LogEntry -Message "Build-JSON Datei nicht gefunden: $($filePath). Setze installedVersion auf null." -Level "WARNING"
            $installedVersion = $null
        }    	
    }

    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localLastWriteTime" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$($ProgramName) ist installiert. Installierte Version: $($installedVersion); Installationsdatei Version: $($localLastWriteTime)" -Level "INFO"

    # Decide whether to install: if installedVersion is available and equals file time -> skip; else install
    if ($installedVersion -and $localLastWriteTime) {
        if ($installedVersion -eq $localLastWriteTime) {
            Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
            $Install = $false
            Write_LogEntry -Message "Installierte Version ist aktuell." -Level "INFO"
        } else {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
            $Install = $true
            Write_LogEntry -Message "Veraltete Version installiert. Update wird gestartet." -Level "WARNING"
        }
    } else {
        # If we cannot determine installedVersion or localLastWriteTime, default to not installing
        # but if InstallationFlag is set we should still run installer (see below)
        Write_LogEntry -Message "Installations- oder Dateiinformationen unvollständig; Install variable bleibt: $($Install)" -Level "DEBUG"
    }
} else {
    Write_LogEntry -Message "Keine Registry-Einträge für $ProgramName gefunden." -Level "WARNING"
    $installedVersion = $null
    $Install = $false
}
Write-Host ""
Write_LogEntry -Message "Installationsprüfung abgeschlossen. Install variable: $($Install); InstallationFlag: $($InstallationFlag)" -Level "DEBUG"

#Install if needed
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte Installationsskript mit Flag." -Level "INFO"
    try {
        & $PSHostPath `
            -NoLogo -NoProfile -ExecutionPolicy Bypass `
            -File "$Serverip\Daten\Prog\InstallationScripts\Installation\jDownload2Install.ps1" `
            -InstallationFlag
        Write_LogEntry -Message "Installationsskript mit Flag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\jDownload2Install.ps1" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Aufruf des Installationsskriptes mit Flag: $($_.Exception.Message)" -Level "ERROR"
    }
}
elseif ($Install -eq $true) {
    Write_LogEntry -Message "Starte Installationsskript (Update) ohne Flag." -Level "INFO"
    try {
        & $PSHostPath `
            -NoLogo -NoProfile -ExecutionPolicy Bypass `
            -File "$Serverip\Daten\Prog\InstallationScripts\Installation\jDownload2Install.ps1"
        Write_LogEntry -Message "Installationsskript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\jDownload2Install.ps1" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Aufruf des Installationsskriptes: $($_.Exception.Message)" -Level "ERROR"
    }
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "DEBUG"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
    Finalize_LogSession | Out-Null
}
# === Ende Logger-Footer ===
