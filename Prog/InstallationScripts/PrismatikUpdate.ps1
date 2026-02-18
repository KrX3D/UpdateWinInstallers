param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Prismatik"
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

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Berechneter Konfigurationspfad: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei geladen: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor Red
    exit 1
}

$wildcardFileName = "Prismatik.unofficial.64bit.Setup*.exe"
Write_LogEntry -Message "Suchmuster für lokale Installationsdatei: $($wildcardFileName)" -Level "DEBUG"

# Lokale Datei bestimmen (letzte im Ordner)
$localFile = Get-ChildItem -Path $InstallationFolder -Filter $wildcardFileName -ErrorAction SilentlyContinue | Select-Object -Last 1
$localFilePath = $null
if ($localFile) {
    $localFilePath = $localFile.FullName
    Write_LogEntry -Message "Lokale Datei gefunden: $($localFilePath)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine lokale Datei gefunden mit Muster $($wildcardFileName) in $($InstallationFolder)" -Level "WARNING"
}

# Lokale Version ermitteln
$localVersion = "0.0.0"
if ($localFilePath -and (Test-Path $localFilePath)) {
    try {
        $fileVersionInfo = Get-ItemProperty -Path $localFilePath -ErrorAction Stop
        $localVersion = $fileVersionInfo.VersionInfo.ProductVersion
        if (-not $localVersion) { $localVersion = "0.0.0" }
        Write_LogEntry -Message "Lokale Dateiversion ermittelt: $($localVersion) für Datei $($localFilePath)" -Level "DEBUG"
    } catch {
        $localVersion = "0.0.0"
        Write_LogEntry -Message "Fehler beim Abrufen der lokalen Dateieigenschaften: $($_)" -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Lokale Datei nicht vorhanden oder Pfad ungültig; setze lokale Version auf $($localVersion)" -Level "DEBUG"
}

# GitHub API konfigurieren
$repoOwner = "psieg"
$repoName  = "Lightpack"
$apiUrl = "https://api.github.com/repos/$repoOwner/$repoName/releases/latest"
Write_LogEntry -Message "GitHub API URL: $($apiUrl)" -Level "DEBUG"

$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) { $headers['Authorization'] = "token $GithubToken" }

# GitHub Release abrufen
$latestRelease = $null
try {
    $latestRelease = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub API Antwort empfangen; Tag: $($latestRelease.tag_name)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen des GitHub Release: $($_)" -Level "ERROR"
    $latestRelease = $null
}

if (-not $latestRelease) {
    Write-Host "Failed to retrieve latest release from GitHub." -ForegroundColor Yellow
    Write_LogEntry -Message "Keine Release-Daten; Abbruch Online-Vergleich." -Level "WARNING"
} else {
    # Version normalisieren: extract first occurrence of numeric version string
    $tagRaw = $latestRelease.tag_name
    $versionMatch = [regex]::Match($tagRaw, '\d+(\.\d+)+')
    if ($versionMatch.Success) {
        $latestVersion = $versionMatch.Value
        Write_LogEntry -Message "Extrahierte Online-Version: $($latestVersion) aus Tag '$tagRaw'." -Level "INFO"
    } else {
        # Fallback: tag komplett verwenden (kann später Probleme bei [version] verursachen)
        $latestVersion = $tagRaw.Trim()
        Write_LogEntry -Message "Konnte keine saubere Versionsnummer extrahieren; verwende TagName: $($latestVersion)" -Level "WARNING"
    }

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -ForegroundColor Cyan
    Write-Host "Online Version: $latestVersion" -ForegroundColor Cyan
    Write-Host ""

    # Versionen vergleichen sicher (versuche [version], sonst string-compare)
    $isNewer = $false
    try {
        if ([version]$latestVersion -gt [version]$localVersion) { $isNewer = $true }
    } catch {
        # nicht parsebar als [version], benutze string-Vergleich (unsicher)
        if ($latestVersion -ne $localVersion) { $isNewer = $true }
    }

    if ($isNewer) {
        Write_LogEntry -Message "Online-Version ist neuer: Online $($latestVersion) > Lokal $($localVersion)" -Level "INFO"

        # Asset-Auswahl: finde exe assets, bevorzuge 64-bit ones
        $assets = $latestRelease.assets
        if (-not $assets -or $assets.Count -eq 0) {
            Write_LogEntry -Message "Keine Assets im Release gefunden." -Level "WARNING"
        } else {
            # Kandidaten: .exe Dateien
            $exeAssets = $assets | Where-Object { $_.name -match '\.exe$' }
            if ($exeAssets.Count -eq 0) {
                Write_LogEntry -Message "Keine EXE-Assets gefunden im Release." -Level "WARNING"
            } else {
                # Versuche zuerst 64-bit/executable with '64' or '64bit' in name
                $preferred = $exeAssets | Where-Object { $_.name -match '64' -or $_.name -match '64bit' } | Sort-Object { $_.name.Length } | Select-Object -First 1
                if (-not $preferred) {
                    # fallback: pick first exe whose name contains 'setup' or 'installer' or 'unofficial'
                    $preferred = $exeAssets | Where-Object { $_.name -match '(setup|installer|unofficial)' } | Sort-Object { $_.name.Length } | Select-Object -First 1
                }
                if (-not $preferred) {
                    # letzter fallback: erstes exe-asset
                    $preferred = $exeAssets | Select-Object -First 1
                }

                if ($preferred) {
                    $downloadLink = $preferred.browser_download_url
                    $downloadFileName = $preferred.name
                    # Zielpfad: im gleichen Verzeichnis wie lokale Datei (oder in InstallationFolder, falls keine lokale Datei)
                    if ($localFilePath) {
                        $targetDir = Split-Path -Path $localFilePath -Parent
                    } else {
                        $targetDir = $InstallationFolder
                    }
                    if (-not (Test-Path $targetDir)) {
                        try { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null } catch { Write_LogEntry -Message "Fehler beim Erstellen Zielverzeichnis $targetDir : $($_)" -Level "ERROR" }
                    }
                    $downloadPath = Join-Path -Path $targetDir -ChildPath $downloadFileName

                    Write_LogEntry -Message "Gewähltes Asset: $($downloadFileName) -> $($downloadLink). Zieldatei: $($downloadPath)" -Level "DEBUG"

                    # Download: WebClient verwenden und Header setzen falls Token vorhanden
                    $webClient = New-Object System.Net.WebClient
                    try {
                        if ($GithubToken) {
                            $webClient.Headers.Add("Authorization", "token $GithubToken")
                            $webClient.Headers.Add("User-Agent", "InstallationScripts/1.0")
                        }
                        Write_LogEntry -Message "Starte Download: $($downloadLink) -> $($downloadPath)" -Level "INFO"
                        $webClient.DownloadFile($downloadLink, $downloadPath)
                        Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "SUCCESS"
                    } catch {
                        Write_LogEntry -Message "Fehler beim Herunterladen $($downloadLink): $($_)" -Level "ERROR"
                        $downloadPath = $null
                    } finally {
                        if ($webClient) { $webClient.Dispose() }
                    }

                    # Nachbearbeitung: ersetzen/verschieben/alten Installer entfernen
                    if ($downloadPath -and (Test-Path $downloadPath)) {
                        try {
                            if ($localFilePath -and (Test-Path $localFilePath)) {
                                Remove-Item -Path $localFilePath -Force -ErrorAction Stop
                                Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFilePath)" -Level "DEBUG"
                            }
                        } catch {
                            Write_LogEntry -Message "Fehler beim Entfernen alter Datei $($localFilePath): $($_)" -Level "WARNING"
                        }

                        Write-Host "$ProgramName wurde aktualisiert: $downloadFileName" -ForegroundColor Green
                        Write_LogEntry -Message "$($ProgramName) Update erfolgreich: $($downloadFileName)" -Level "SUCCESS"
                    } else {
                        Write_LogEntry -Message "Download fehlgeschlagen oder Datei nicht vorhanden nach Download." -Level "ERROR"
                    }
                } else {
                    Write_LogEntry -Message "Kein geeignetes Asset ausgewählt (preferred ist leer)." -Level "WARNING"
                }
            }
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write_LogEntry -Message "Kein Online Update verfügbar. Online: $($latestVersion); Lokal: $($localVersion)" -Level "INFO"
    }
}

# --- Installationsprüfung / Aufruf Installation Script falls benötigt ---
# Bestimme erneut lokale Datei (aktualisiert)
$localFile2 = Get-ChildItem -Path $InstallationFolder -Filter $wildcardFileName -ErrorAction SilentlyContinue | Select-Object -Last 1
if ($localFile2) {
    try { $localVersion = (Get-ItemProperty -Path $localFile2.FullName).VersionInfo.ProductVersion } catch { $localVersion = "0.0.0" }
    Write_LogEntry -Message "Erneut lokale Datei bestimmt: $($localFile2.FullName); Version: $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine lokale Datei beim erneuten Check gefunden." -Level "WARNING"
}

# Registry check (wie vorher)
$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

$Install = $false
if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    if ($installedVersion) {
        try {
            if ([version]$installedVersion -lt [version]$localVersion) {
                $Install = $true
            }
        } catch {
            # fallback: string compare
            if ($installedVersion -ne $localVersion) { $Install = $true }
        }
    }
}

# Install/Update ausführen
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte Installationsskript mit Flag." -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\PrismatikInstallation.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Installationsskript mit Flag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrismatikInstallation.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte Installationsskript (Update) ohne Flag: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrismatikInstallation.ps1" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\PrismatikInstallation.ps1"
    Write_LogEntry -Message "Installationsskript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrismatikInstallation.ps1" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"
# === Logger-Footer ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===
