param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinMerge"
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
    exit 1
}

# Define the path to the local WinMerge installer
$localInstallerPath = Join-Path -Path $InstallationFolder -ChildPath "WinMerge-*.exe"
Write_LogEntry -Message "Suche lokale Installer mit Pattern: $($localInstallerPath)" -Level "DEBUG"

# Get the local installer file
$localInstaller = Get-ChildItem -Path $localInstallerPath -ErrorAction SilentlyContinue | Select-Object -First 1
if ($localInstaller) {
    Write_LogEntry -Message "Gefundene lokale Installationsdatei: $($localInstaller.FullName)" -Level "DEBUG"

    # Get the file version from the local installer properties
    try {
        $localVersion = (Get-ItemProperty -LiteralPath $localInstaller.FullName -ErrorAction Stop).VersionInfo.FileVersion
        Write_LogEntry -Message "Ermittelte lokale Dateiversion: $($localVersion) aus Datei $($localInstaller.FullName)" -Level "INFO"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Dateiversion für $($localInstaller.FullName): $($_)" -Level "ERROR"
        $localVersion = $null
    }

    # Prepare GitHub API headers (use token if available)
    $releasesURL = "https://api.github.com/repos/WinMerge/winmerge/releases"
    Write_LogEntry -Message "Rufe GitHub Releases ab: $($releasesURL)" -Level "DEBUG"

    $headers = @{
        'User-Agent' = 'InstallationScripts/1.0'
        'Accept'     = 'application/vnd.github.v3+json'
    }
    
    # Use $GithubToken from config if present
    if ($GithubToken) {
        $headers['Authorization'] = "token $GithubToken"
        Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfragen." -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Kein GitHub Token gefunden; benutze anonyme API-Anfragen (niedrigere Rate-Limits)." -Level "DEBUG"
    }

    # Request releases from GitHub with headers
    $latestRelease = $null
    try {
        $latestReleaseRaw = Invoke-RestMethod -Uri $releasesURL -Headers $headers -ErrorAction Stop
        Write_LogEntry -Message "GitHub Releases abgerufen; Elementanzahl: $($latestReleaseRaw.Count)" -Level "DEBUG"
        
        # Filter out prereleases and unwanted tags (ShellExtension, Merge7z, etc.)
        # Only keep releases that start with 'v' followed by digits (main WinMerge releases)
        $filtered = $latestReleaseRaw | Where-Object { 
            $_.prerelease -eq $false -and 
            $_.tag_name -match '^v?\d+\.\d+' -and
            $_.tag_name -notmatch '^(Merge7z|ShellExtension|WinIMerge)'
        }
        
        if ($filtered -and $filtered.Count -gt 0) {
            $latestRelease = $filtered | Sort-Object {[datetime]$_.published_at} -Descending | Select-Object -First 1
            Write_LogEntry -Message "Gefilterte Releases gefunden; ausgewählte neueste Release: $($latestRelease.tag_name)" -Level "DEBUG"
        } else {
            Write_LogEntry -Message "Keine passenden Releases nach Filter gefunden." -Level "WARNING"
            $latestRelease = $null
        }
    } catch {
        # Try to extract a helpful message if GitHub returned a JSON error (rate limit etc.)
        $errMsg = $_.Exception.Message
        try {
            if ($_.Exception.Response -ne $null) {
                $respStream = $_.Exception.Response.GetResponseStream()
                $reader = [System.IO.StreamReader]::new($respStream)
                $respBody = $reader.ReadToEnd()
                $reader.Close()
                $respStream.Close()
                if ($respBody) {
                    try {
                        $errJson = $respBody | ConvertFrom-Json -ErrorAction Stop
                        if ($errJson.message) { $errMsg = $errJson.message }
                    } catch {
                        # ignore parse error
                    }
                }
            }
        } catch {
            # ignore
        }

        if ($errMsg -match '(rate limit|rate_limit|rate limit exceeded|API rate limit|403)') {
            Write_LogEntry -Message "GitHub API Rate-Limit / Zugriff verweigert erkannt: $($errMsg)" -Level "WARNING"
            Write_LogEntry -Message "Hinweis: Lege einen GitHub-PAT in PowerShellVariables.ps1 als `$GithubToken` ab, um höhere Limits zu erhalten." -Level "INFO"
        } else {
            Write_LogEntry -Message "Fehler beim Abrufen der GitHub Releases: $($errMsg)" -Level "ERROR"
        }
        $latestRelease = $null
    }

    if ($latestRelease) {
        # Normalize the tag (strip leading v or 'version_' if present)
        $latestVersionRaw = $latestRelease.tag_name
        $latestVersion = $latestVersionRaw -replace '^v','' -replace '^version_',''
        Write_LogEntry -Message "Ermittelte Online-Version (raw/tag): $($latestVersionRaw) -> normalized: $($latestVersion)" -Level "INFO"

        # determine best asset (prefer x64 Setup asset, exclude ARM64)
        $asset = $null
        try {
            if ($latestRelease.assets -and $latestRelease.assets.Count -gt 0) {
                # Try to find x64 setup exe, explicitly excluding ARM64
                $asset = $latestRelease.assets | Where-Object {
                    $_.name -match 'x64' -and 
                    $_.name -notmatch 'ARM64' -and 
                    $_.name -notmatch 'PerUser' -and 
                    $_.name -match '\.exe$' -and
                    ($_.name -match 'setup' -or $_.name -match 'Setup')
                } | Select-Object -First 1

                # If not found with strict criteria, try broader x64 match
                if (-not $asset) {
                    $asset = $latestRelease.assets | Where-Object { 
                        $_.name -match 'x64' -and 
                        $_.name -notmatch 'ARM64' -and 
                        $_.name -notmatch 'PerUser' -and 
                        $_.name -match '\.exe$' 
                    } | Select-Object -First 1
                }

                # Last resort: any exe that's not ARM64
                if (-not $asset) {
                    $asset = $latestRelease.assets | Where-Object { 
                        $_.name -match '\.exe$' -and 
                        $_.name -notmatch 'PerUser' -and 
                        $_.name -notmatch 'ARM64' 
                    } | Select-Object -First 1
                }
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Auswählen des Release-Assets: $($_)" -Level "WARNING"
            $asset = $null
        }

        # Build download URL and filename
        if ($asset -ne $null) {
            $downloadURL = $asset.browser_download_url
            $filename = $asset.name
            Write_LogEntry -Message "Download-Asset gewählt: $($filename); URL: $($downloadURL)" -Level "DEBUG"
        } else {
            # fallback constructed URL (may fail if releases naming differs)
            $downloadURL = "https://github.com/WinMerge/winmerge/releases/download/v$latestVersion/WinMerge-$latestVersion-x64-Setup.exe"
            $filename = [System.IO.Path]::GetFileName($downloadURL)
            Write_LogEntry -Message "Kein Asset gefunden; konstruiere Fallback-Download-URL: $($downloadURL)" -Level "WARNING"
        }

        Write-Host ""
        Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
        Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
        Write-Host ""

        # safe version parsing for comparison
        $canParseLocal = $false; $canParseOnline = $false
        try { $localVobj = if ($localVersion) { [version]($localVersion -replace '[^\d\.]','') } else { $null } ; if ($localVobj) { $canParseLocal = $true } } catch { $canParseLocal = $false }
        try { $onlineVobj = if ($latestVersion) { [version]($latestVersion -replace '[^\d\.]','') } else { $null } ; if ($onlineVobj) { $canParseOnline = $true } } catch { $canParseOnline = $false }

        $needUpdate = $false
        if ($canParseLocal -and $canParseOnline) {
            if ($onlineVobj -gt $localVobj) { $needUpdate = $true }
        } else {
            # fallback string compare if versions couldn't be parsed reliably
            if ($latestVersion -and $localVersion) {
                if ($latestVersion -ne $localVersion) { $needUpdate = $true }
            } elseif ($latestVersion -and -not $localVersion) {
                $needUpdate = $true
            }
        }

        if ($needUpdate) {
            $downloadPath = Join-Path -Path $InstallationFolder -ChildPath $filename
            Write_LogEntry -Message "Update verfügbar: Local $($localVersion) < Online $($latestVersion). Downloadpfad: $($downloadPath)" -Level "INFO"

            try {
                # Download with WebClient to file; add headers if token present
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("user-agent", $headers['User-Agent'])
                if ($headers.ContainsKey('Authorization')) {
                    # webclient expects "Authorization" as string header
                    $webClient.Headers.Add("Authorization", $headers['Authorization'])
                }
                Write_LogEntry -Message "Starte Download von $($downloadURL) -> $($downloadPath)" -Level "INFO"
                [void](Invoke-DownloadFile -Url $downloadURL -OutFile $downloadPath)
                Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "Fehler beim Herunterladen $($downloadURL): $($_)" -Level "ERROR"
            } finally {
                if ($webClient) { $webClient.Dispose() }
            }

            # Check if the file was completely downloaded
            if (Test-Path $downloadPath) {
                try {
                    # Remove the old installer (best-effort)
                    Remove-Item -Path $localInstaller.FullName -Force -ErrorAction SilentlyContinue
                    Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localInstaller.FullName)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($localInstaller.FullName): $($_)" -Level "WARNING"
                }

                Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                Write_LogEntry -Message "$($ProgramName) Update erfolgreich: $($latestVersion)" -Level "SUCCESS"
            } else {
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "red"
                Write_LogEntry -Message "Download fehlgeschlagen für $($downloadPath)" -Level "ERROR"
            }
        } else {
            Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
            Write_LogEntry -Message "Keine neuere Version verfügbar: Local $($localVersion) >= Online $($latestVersion)" -Level "INFO"
        }
    } else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Keine Online-Version ermittelt; Update-Vergleich übersprungen." -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Lokaler WinMerge-Installer nicht gefunden mit Pattern: $($localInstallerPath)" -Level "WARNING"
}

Write-Host ""

#Check Installed Version / Install if needed
$FoundFile = Get-ChildItem -Path $localInstallerPath -ErrorAction SilentlyContinue | Select-Object -First 1
if ($FoundFile) {
    Write_LogEntry -Message "Gefundene Datei für Install-Check: $($FoundFile.FullName)" -Level "DEBUG"
    try {
        $localVersion = (Get-ItemProperty -Path $FoundFile.FullName -ErrorAction Stop).VersionInfo.FileVersion
        Write_LogEntry -Message "Installationsdatei Version ermittelt: $($localVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der Installationsdatei-Version: $($_)" -Level "WARNING"
    }

    $InstallationFileName = $FoundFile.Name
    $localInstaller = Join-Path -Path $InstallationFolder -ChildPath $InstallationFileName
    try {
        $localVersion = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).VersionInfo.ProductVersion
        Write_LogEntry -Message "Produktversion der Installationsdatei: $($localVersion) (Pfad: $($localInstaller))" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der Produktversion für $($localInstaller): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Keine Datei für Install-Check gefunden mit Pattern: $($localInstallerPath)" -Level "DEBUG"
}

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Prüfe Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registry-Pfad gefunden: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    } else {
        Write_LogEntry -Message "Registry-Pfad existiert nicht: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path) {
    try {
        $installedVersion = $Path.DisplayVersion | Select-Object -First 1
        Write_LogEntry -Message "Gefundene installierte Version aus Registry: $($installedVersion)" -Level "INFO"
    } catch {
        Write_LogEntry -Message "Fehler beim Auslesen der installierten Version aus Registry: $($_)" -Level "ERROR"
        $installedVersion = $null
    }

    if ($null -ne $installedVersion) {
        Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
        Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
        Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
        Write_LogEntry -Message "Installationsstatus: installiert; InstalledVersion: $($installedVersion); LocalVersion: $($localVersion)" -Level "INFO"
    
        try {
            if ([version]$installedVersion -lt [version]$localVersion) {
                Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
                $Install = $true
                Write_LogEntry -Message "Install erforderlich: Installed $($installedVersion) < Local $($localVersion)" -Level "INFO"
            } elseif ([version]$installedVersion -eq [version]$localVersion) {
                Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
                $Install = $false
                Write_LogEntry -Message "Install nicht erforderlich: InstalledVersion == LocalVersion ($($localVersion))" -Level "INFO"
            } else {
                $Install = $false
                Write_LogEntry -Message "Install nicht ausgeführt: InstalledVersion ($($installedVersion)) > LocalVersion ($($localVersion))" -Level "WARNING"
            }
        } catch {
            # If version parsing fails, fallback to conservative decision: no install
            $Install = $false
            Write_LogEntry -Message "Fehler beim Vergleichen der Versionen (Parsing). Install=false. Fehler: $($_)" -Level "WARNING"
        }
    } else {
        $Install = $false
        Write_LogEntry -Message "Keine installierte Version aus Registry ermittelt; Install=false" -Level "DEBUG"
    }
} else {
    $Install = $false
    Write_LogEntry -Message "Keine Registry-Einträge für $($ProgramName) gefunden; Install=false" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Installationsprüfung abgeschlossen. Install variable: $($Install)" -Level "DEBUG"

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "Starte externes Installationsskript mit -InstallationFlag" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\WinMergeInstallation.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WinMergeInstallation.ps1 (mit -InstallationFlag)" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte externes Installationsskript (Install=true)" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\WinMergeInstallation.ps1"
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WinMergeInstallation.ps1" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht" -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===

