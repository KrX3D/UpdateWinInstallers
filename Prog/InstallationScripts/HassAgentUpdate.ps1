param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Hass.Agent"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PSScriptRoot: $($PSScriptRoot)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

# Define the directory path and file wildcard
$InstallationFolder = "$NetworkShareDaten\Projekte\Smart_Home\HASS_Agent"
#$fileWildcard = "HASS.Agent.Installer_*.exe"
$fileWildcard = "HASS.Agent.Installer.exe"
$IncludeBeta = $true

Write_LogEntry -Message "InstallationFolder: $($InstallationFolder); FileWildcard: $($fileWildcard); IncludeBeta: $($IncludeBeta)" -Level "DEBUG"

# Get the latest local file path matching the wildcard (defensive)
$localFilePath = $null
try {
    if (Test-Path -Path $InstallationFolder) {
        $localFileObj = Get-ChildItem -Path $InstallationFolder -Filter $fileWildcard -File -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($localFileObj) { $localFilePath = $localFileObj.FullName }
    } else {
        Write_LogEntry -Message "Installationsordner existiert nicht: $($InstallationFolder)" -Level "WARNING"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Suchen lokaler Dateien: $($_)" -Level "WARNING"
}
Write_LogEntry -Message "Gefundener lokaler Datei-Pfad: $($localFilePath)" -Level "DEBUG"

# Get the version number from the file properties (defensive)
$localVersion = "0.0.0"
$localBetaVersion = $null
if ($localFilePath -and (Test-Path -Path $localFilePath)) {
    try {
        $fileVersionInfo = (Get-Item -LiteralPath $localFilePath -ErrorAction Stop).VersionInfo
        $localVersionRaw = ($fileVersionInfo.ProductVersion -split ' ')[0]
        $localBetaVersion = if ($localVersionRaw -match '-beta(\d+)') { $Matches[1] } else { $null }
        $localVersion = $localVersionRaw -replace '-beta\d+', ''
        Write_LogEntry -Message "Produkt-Version Info aus Datei geladen: $($localFilePath)" -Level "DEBUG"
        Write_LogEntry -Message "Lokale Version: $($localVersion); Lokale Beta: $($localBetaVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Produktversion für $localFilePath : $($_)" -Level "WARNING"
        $localVersion = "0.0.0"
        $localBetaVersion = $null
    }
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei gefunden; setze lokale Version auf $($localVersion)" -Level "DEBUG"
}

# Prepare GitHub API URL and headers (use token if available)
$repository = "hass-agent/HASS.Agent"
if ($IncludeBeta) {
    $apiUrl = "https://api.github.com/repos/$repository/releases"
} else {
    $apiUrl = "https://api.github.com/repos/$repository/releases/latest"
}
Write_LogEntry -Message "GitHub API URL: $($apiUrl); Repository: $($repository)" -Level "INFO"

$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein GitHub Token gefunden; benutze anonyme API-Anfragen (niedrigere Rate-limits)." -Level "DEBUG"
}

# Fetch release(s) from GitHub
$release = $null
try {
    $release = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub API abgefragt; Rohantwort erhalten." -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen der GitHub API: $($_)" -Level "ERROR"
    $release = $null
}

# Normalize and pick the latest release (handle both single and array responses)
$chosenRelease = $null
try {
    if ($release -eq $null) {
        Write_LogEntry -Message "Keine Release-Informationen empfangen; überspringe Online-Check." -Level "WARNING"
    } else {
        if ($IncludeBeta) {
            # $release is expected to be an array; pick the newest (by published_at or fallback by tag_name version)
            if ($release -is [System.Array]) {
                $candidates = $release | Where-Object { -not $_.draft } # allow prereleases; we don't filter out prereleases here if IncludeBeta is true
                if ($candidates.Count -eq 0) { $candidates = $release } # fallback to all
                $chosenRelease = $candidates | Sort-Object @{ Expression = { if ($_.published_at) { [datetime]$_.published_at } else { [datetime]::MinValue } } } -Descending | Select-Object -First 1
            } else {
                # single object returned unexpectedly; use it
                $chosenRelease = $release
            }
        } else {
            # latest (single) response
            $chosenRelease = $release
        }
    }
} catch {
    Write_LogEntry -Message "Fehler beim Auswählen des Releases: $($_)" -Level "ERROR"
    $chosenRelease = $null
}

if (-not $chosenRelease) {
    Write_LogEntry -Message "Kein Release ausgewählt; Ende Online-Prüfung." -Level "WARNING"
} else {
    # Extract web version and beta index if present
    $rawTag = $chosenRelease.tag_name
    $webVersionRaw = if ($rawTag) { ($rawTag -split ' ')[0] } else { $chosenRelease.name }
    $webBetaVersion = $null
    if ($webVersionRaw -match '-beta(\d+)') {
        $webBetaVersion = $Matches[1]
    }
    $webVersion = $webVersionRaw -replace '-beta\d+',''
    Write_LogEntry -Message "Ermittelte Online-Version: $($webVersion); Online Beta: $($webBetaVersion); Release Tag: $($rawTag)" -Level "INFO"

    Write-Host ""
    if ($localBetaVersion) {
        Write-Host "Lokale Version: $localVersion Beta: $localBetaVersion" -foregroundcolor "Cyan"
    } else {
        Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    }
    if ($webBetaVersion) {
        Write-Host "Online Version: $webVersion Beta: $webBetaVersion" -foregroundcolor "Cyan"
    } else {
        Write-Host "Online Version: $webVersion" -foregroundcolor "Cyan"
    }
    Write-Host ""

    # Decide if update is needed (use [version] where possible)
    $needUpdate = $false
    try {
        if ([version]$webVersion -gt [version]$localVersion) { $needUpdate = $true }
        elseif ([version]$webVersion -eq [version]$localVersion) {
            if ($webBetaVersion -and $localBetaVersion) {
                if ([int]$webBetaVersion -gt [int]$localBetaVersion) { $needUpdate = $true }
            } elseif ($webBetaVersion -and -not $localBetaVersion) {
                # online is beta for same base, consider update if you want beta over stable (your logic may vary)
                $needUpdate = $true
            } else {
                $needUpdate = $false
            }
        } else { $needUpdate = $false }
    } catch {
        # fallback string comparison if version parse fails
        if ($webVersion -ne $localVersion) { $needUpdate = $true }
    }

    if ($needUpdate) {
        $updateMessage = "Update erkannt: Online $($webVersion) (beta $($webBetaVersion)) > Lokal $($localVersion) (beta $($localBetaVersion))."
        Write_LogEntry -Message $updateMessage -Level "INFO"

        # Select asset: prefer installer with typical naming, else any .exe
        $asset = $null
        try {
            if ($chosenRelease.assets -and $chosenRelease.assets.Count -gt 0) {
                # first try specific pattern
                $asset = $chosenRelease.assets | Where-Object { $_.name -match 'Installer' -and $_.name -match '\.exe$' } | Select-Object -First 1
                if (-not $asset) {
                    # fallback to any exe asset
                    $asset = $chosenRelease.assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1
                }
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Auswählen des Assets: $($_)" -Level "WARNING"
            $asset = $null
        }

        if (-not $asset) {
            Write_LogEntry -Message "Kein passendes Asset (Installer .exe) im Release gefunden; Update übersprungen." -Level "WARNING"
        } else {
            $downloadLink = $asset.browser_download_url
            $assetName = $asset.name
            Write_LogEntry -Message "Download-Asset gefunden: $($assetName) -> $($downloadLink)" -Level "DEBUG"

            # Construct target filename and do a safe download (temp .part -> move)
            $newFilePath = Join-Path -Path $InstallationFolder -ChildPath $assetName
            $tempFile = "$newFilePath.part"

            Write_LogEntry -Message "Starte Download von $($downloadLink) nach (temp) $($tempFile)" -Level "INFO"
            $downloadOk = $false
            try {
                $wc = New-Object System.Net.WebClient
                if ($headers.ContainsKey('User-Agent')) { $wc.Headers.Add('User-Agent', $headers['User-Agent']) }
                if ($headers.ContainsKey('Authorization')) { $wc.Headers.Add('Authorization', $headers['Authorization']) }
                [void](Invoke-DownloadFile -Url $downloadLink -OutFile $tempFile)
                $wc.Dispose()
                $downloadOk = Test-Path -Path $tempFile
                Write_LogEntry -Message "Temp-Download abgeschlossen: $($tempFile) (exists: $($downloadOk))" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Herunterladen des Assets: $($_)" -Level "ERROR"
                try { if ($wc) { $wc.Dispose() } } catch {}
                $downloadOk = $false
            }

            if ($downloadOk) {
                try {
                    Move-Item -Path $tempFile -Destination $newFilePath -Force -ErrorAction Stop
                    Write_LogEntry -Message "Temp-Datei verschoben nach: $($newFilePath)" -Level "DEBUG"

                    # remove old installer (if different)
                    if ($localFilePath -and (Test-Path -Path $localFilePath) -and ($localFilePath -ne $newFilePath)) {
                        try {
                            Remove-Item -Path $localFilePath -Force -ErrorAction Stop
                            Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFilePath)" -Level "DEBUG"
                        } catch {
                            Write_LogEntry -Message "Fehler beim Entfernen alter Installationsdatei: $($_)" -Level "WARNING"
                        }
                    }

                    Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                    $successMessage = "$($ProgramName) wurde aktualisiert auf $($webVersion) (beta: $($webBetaVersion)). Neue Datei: $($newFilePath)"
                    Write_LogEntry -Message $successMessage -Level "SUCCESS"

                    # update local variables for subsequent checks
                    $localFilePath = $newFilePath
                    try {
                        $fileVersionInfo = (Get-Item -LiteralPath $localFilePath).VersionInfo
                        $localVersionRaw = ($fileVersionInfo.ProductVersion -split ' ')[0]
                        $localBetaVersion = if ($localVersionRaw -match '-beta(\d+)') { $Matches[1] } else { $null }
                        $localVersion = $localVersionRaw -replace '-beta\d+',''
                        Write_LogEntry -Message "Nach Update: lokale Version ist nun $($localVersion) (beta $($localBetaVersion))." -Level "DEBUG"
                    } catch {
                        Write_LogEntry -Message "Nach Update: Fehler beim Ermitteln der lokalen Version: $($_)" -Level "WARNING"
                    }

                } catch {
                    Write_LogEntry -Message "Fehler beim Finalisieren des Downloads (Move/Remove): $($_)" -Level "ERROR"
                    Write-Host "Download konnte nicht finalisiert werden." -ForegroundColor Red
                }
            } else {
                Write_LogEntry -Message "Download nicht erfolgreich; temp-Datei nicht vorhanden: $($tempFile)" -Level "ERROR"
                try { if (Test-Path $tempFile) { Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue } } catch {}
            }
        }
    } else {
        Write_LogEntry -Message "Kein Online Update verfügbar. $($ProgramName) ist aktuell (Online: $($webVersion); Lokal: $($localVersion))." -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    }
}

Write-Host ""
Write_LogEntry -Message "Erneute Bestimmung lokaler Datei und Version für Installationsprüfung." -Level "DEBUG"

#Check Installed Version / Install if needed
$localFilePath = $null
try {
    if (Test-Path -Path $InstallationFolder) {
        $localFileObj = Get-ChildItem -Path $InstallationFolder -Filter $fileWildcard -File -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($localFileObj) { $localFilePath = $localFileObj.FullName }
    }
} catch {
    Write_LogEntry -Message "Fehler beim Auffinden lokaler Datei: $($_)" -Level "WARNING"
}
Write_LogEntry -Message "Lokale Datei für Installationsprüfung: $($localFilePath)" -Level "DEBUG"

if ($localFilePath -and (Test-Path -Path $localFilePath)) {
    try {
        $fileVersionInfo = (Get-Item -LiteralPath $localFilePath).VersionInfo
        $localVersionRaw = ($fileVersionInfo.ProductVersion -split ' ')[0]
        $localBetaVersion = if ($localVersionRaw -match '-beta(\d+)') { $Matches[1] } else { $null }
        $localVersion = $localVersionRaw -replace '-beta\d+',''
        Write_LogEntry -Message "Ermittelte lokale Version (erneut): $($localVersion); Beta: $($localBetaVersion)" -Level "DEBUG"
    } catch {
        $localVersion = "0.0.0"
        $localBetaVersion = $null
        Write_LogEntry -Message "Fehler beim Lesen der lokalen Version für Installationsprüfung: $($_)" -Level "WARNING"
    }
} else {
    $localVersion = "0.0.0"
    $localBetaVersion = $null
    Write_LogEntry -Message "Keine lokale Datei für Installationsprüfung gefunden; setze lokale Version auf $($localVersion)" -Level "DEBUG"
}

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Zu prüfende Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
    $installedVersionRaw = $Path.DisplayVersion | Select-Object -First 1
    $installedBetaVersion = if ($installedVersionRaw -match '-beta(\d+)') { $Matches[1] } else { $null }
    $installedVersion = if ($installedVersionRaw) { $installedVersionRaw -replace '-beta\d+','' } else { $null }

    Write_LogEntry -Message "Gefundene installierte Version: $($installedVersion); Installed Beta: $($installedBetaVersion)" -Level "INFO"

    if ($installedVersion) {
        if ($installedVersion -lt $localVersion -or ($installedVersion -eq $localVersion -and $installedBetaVersion -lt $localBetaVersion)) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
            $versionComparisonMessage = "Installierte Version $($installedVersion) mit beta $($installedBetaVersion) ist aelter als lokale Datei $($localVersion) mit beta $($localBetaVersion). Markiere Installation."
            Write_LogEntry -Message $versionComparisonMessage -Level "INFO"
            $Install = $true
        } elseif ($installedVersion -eq $localVersion -and $installedBetaVersion -eq $localBetaVersion) {
            Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
            Write_LogEntry -Message "Installierte Version entspricht Installationsdatei: $($installedVersion) (beta: $($installedBetaVersion))." -Level "DEBUG"
            $Install = $false
        } elseif ($installedBetaVersion -and -not $localBetaVersion -and $installedVersion -eq $localVersion) {
            Write-Host "        Stable Version von $ProgramName verfügbar. Update wird gestartet." -foregroundcolor "magenta"
            Write_LogEntry -Message "Installed is beta and local is stable; schedule upgrade to stable." -Level "INFO"
            $Install = $true
        } else {
            Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer oder gleich; kein Update nötig." -Level "DEBUG"
            $Install = $false
        }
    } else {
        Write_LogEntry -Message "Keine klare installierte Version in Registry gefunden; Install=false" -Level "DEBUG"
        $Install = $false
    }
} else {
    Write_LogEntry -Message "$($ProgramName) nicht in der Registry gefunden. Install-Flag auf $($false) gesetzt." -Level "INFO"
    $Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1 mit Parameter -InstallationFlag" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Install Flag true: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1"
    Write_LogEntry -Message "Externes Installations-Skript für Hass.Agent aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1" -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
