param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Microsoft Windows Desktop Runtime"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

# Define the directory path and file wildcard
$InstallationFolder = "$InstallationFolder\ImageGlass"
$fileWildcard = "windowsdesktop-runtime-*-win-x64.exe"
Write_LogEntry -Message "InstallationFolder: $($InstallationFolder); FileWildcard: $($fileWildcard)" -Level "DEBUG"

# Get the latest local file path matching the wildcard
$localFilePath = $null
try {
    $localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $fileWildcard -ErrorAction Stop | Select-Object -Last 1 -ExpandProperty FullName
    if ($localFilePath) {
        Write_LogEntry -Message "Lokale Installationsdatei gefunden: $($localFilePath)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Keine lokale Installationsdatei gefunden im Pfad: $($InstallationFolder) mit Filter $($fileWildcard)" -Level "WARNING"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Suchen lokaler Dateien: $($_)" -Level "WARNING"
}

# Get the version number from the file properties
try {
    if ($localFilePath -and (Test-Path -Path $localFilePath)) {
        $fileVersionInfo = (Get-Item -LiteralPath $localFilePath -ErrorAction Stop).VersionInfo
        # keep only first three components for comparison
        $localVersion = ($fileVersionInfo.FileVersion -split "\." | Select-Object -First 3) -join "."
        Write_LogEntry -Message "Ermittelte lokale Dateiversion: $($localVersion) aus Datei $($localFilePath)" -Level "INFO"
    } else {
        $localVersion = "0.0.0"
        Write_LogEntry -Message "Lokale Datei nicht vorhanden. Setze lokale Version auf $($localVersion)" -Level "DEBUG"
    }
} catch {
    $localVersion = "0.0.0"
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Dateiversion für $($localFilePath): $($_)" -Level "ERROR"
}

# Define the repository and API URL
$repository = "dotnet/windowsdesktop"
$apiUrl = "https://api.github.com/repos/$repository/releases"
$targetMajorVersion = "8.0"
Write_LogEntry -Message "GitHub API URL: $($apiUrl); TargetMajorVersion: $($targetMajorVersion)" -Level "DEBUG"

# Prepare GitHub API headers, use token if configured
$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }

if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein GitHub Token gefunden; benutze anonyme API-Anfragen (niedrigere Rate-Limits)." -Level "DEBUG"
}

# Fetch releases from GitHub (with token if present)
$releasesResponse = $null
try {
    $releasesResponse = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub Releases abgerufen; Anzahl Releases: $($releasesResponse.Count)" -Level "DEBUG"
} catch {
    $errMsg = $_.Exception.Message
    # try to parse response body for nicer message
    try {
        if ($_.Exception.Response) {
            $rs = $_.Exception.Response.GetResponseStream()
            $rr = [System.IO.StreamReader]::new($rs)
            $rb = $rr.ReadToEnd(); $rr.Close(); $rs.Close()
            try { $rbj = $rb | ConvertFrom-Json -ErrorAction Stop; if ($rbj.message) { $errMsg = $rbj.message } } catch {}
        }
    } catch {}
    Write_LogEntry -Message "Fehler beim Abruf der Releases von GitHub $($apiUrl): $($errMsg)" -Level "ERROR"
    $releasesResponse = $null
}

# Extract the latest release for the specified major version
$latestRelease = $null
if ($releasesResponse) {
    try {
        $latestRelease = $releasesResponse |
            Where-Object { $_.tag_name -like "v$targetMajorVersion.*" -and -not $_.prerelease -and -not $_.draft } |
            Sort-Object { 
                # try to sort by published_at if available; else by parsed version
                if ($_.published_at) { [datetime]$_.published_at } else { [version]($_.tag_name -replace '^v','') }
            } -Descending |
            Select-Object -First 1

        if ($latestRelease) {
            Write_LogEntry -Message "Gefundener neuester Release-Tag: $($latestRelease.tag_name)" -Level "DEBUG"
        } else {
            Write_LogEntry -Message "Kein passender Release für MajorVersion $($targetMajorVersion) gefunden" -Level "WARNING"
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Filtern/Sortieren der Releases: $($_)" -Level "ERROR"
        $latestRelease = $null
    }
} else {
    Write_LogEntry -Message "Keine ReleasesResponse vorhanden; überspringe Release-Analyse" -Level "ERROR"
}

# If we have a release, try to find a downloadable installer asset or fallback to parsing dotnet page
if ($latestRelease) {
    $webVersion = $latestRelease.tag_name -replace "^v", ""
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $webVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Lokale Version: $($localVersion); Online Version: $($webVersion)" -Level "INFO"
    Write-Host ""

    # First: try to find an asset on the release that matches the windowsdesktop runtime installer
    $directDownloadUrl = $null
    $assetName = $null
    try {
        if ($latestRelease.assets -and $latestRelease.assets.Count -gt 0) {
            # prefer asset names containing windowsdesktop, runtime, win, x64, exe
            $asset = $latestRelease.assets | Where-Object {
                ($_.name -match 'windowsdesktop' -or $_.name -match 'desktop') -and
                ($_.name -match 'runtime' -or $_.name -match 'Runtime') -and
                ($_.name -match 'win' -or $_.name -match 'x64') -and
                ($_.name -match '\.exe$')
            } | Select-Object -First 1

            if (-not $asset) {
                # try a more relaxed match: any exe with 'win' or 'x64'
                $asset = $latestRelease.assets | Where-Object { ($_.name -match '\.exe$') -and ($_.name -match 'win|x64') } | Select-Object -First 1
            }

            if ($asset) {
                $directDownloadUrl = $asset.browser_download_url
                $assetName = $asset.name
                Write_LogEntry -Message "Gefundenes Release-Asset: $($assetName) -> $($directDownloadUrl)" -Level "DEBUG"
            } else {
                Write_LogEntry -Message "Kein passendes Release-Asset für Installer gefunden; versuche Fallback (dotnet-Seite)." -Level "WARNING"
            }
        } else {
            Write_LogEntry -Message "Release hat keine Assets; versuche Fallback (dotnet-Seite)." -Level "WARNING"
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Prüfen von Release-Assets: $($_)" -Level "WARNING"
    }

    # Fallback: try to obtain direct link from official dotnet download 'thank-you' page
    if (-not $directDownloadUrl) {
        $downloadLink = "https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-$webVersion-windows-x64-installer"
        Write_LogEntry -Message "Fallback: Rufe Dotnet Download-Seite ab: $($downloadLink)" -Level "DEBUG"
        $response = $null
        try {
            #$response = Invoke-WebRequest -Uri $downloadLink -Headers @{ 'User-Agent' = 'InstallationScripts/1.0' } -ErrorAction Stop
			$response = Invoke-WebRequest -Uri $downloadLink -Headers @{ 'User-Agent' = 'InstallationScripts/1.0' } -UseBasicParsing -ErrorAction Stop
            Write_LogEntry -Message "Download-Seite abgerufen; StatusCode: $($response.StatusCode)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Abrufen der Dotnet-Downloadseite $($downloadLink): $($_)" -Level "ERROR"
            $response = $null
        }

        if ($response) {
            # Try to find a link element with id="directLink" (existing logic) or any exe link that contains windowsdesktop-runtime & win-x64
            try {
                $linkObj = $null
                if ($response.Links) {
                    $linkObj = $response.Links | Where-Object { $_.id -eq 'directLink' } | Select-Object -First 1
                }
                if (-not $linkObj) {
                    # regex search in raw content for a likely exe URL (common pattern)
                    $pattern = 'https?:\/\/[^"''`>]+windowsdesktop[^"''`>]*win[^"''`>]*x64[^"''`>]*\.exe'
                    $match = [regex]::Match($response.Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

                    if ($match.Success) {
                        $directDownloadUrl = $match.Value
                        Write_LogEntry -Message "Gefundene Download-URL via Regex auf Seite: $($directDownloadUrl)" -Level "DEBUG"
                    } else {
                        Write_LogEntry -Message "Kein direkter .exe-Link auf der Dotnet-Seite gefunden (Regex)." -Level "WARNING"
                    }
                } else {
                    $directDownloadUrl = $linkObj.href
                    Write_LogEntry -Message "Direkter Download-Link (id=directLink) gefunden: $($directDownloadUrl)" -Level "DEBUG"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Extrahieren des direkten Download-Links aus der Dotnet-Seite: $($_)" -Level "WARNING"
            }
        }
    }

    # decide if we actually need to download (don't download if versions are equal)
    $needDownload = $false
    try {
        if (-not $localVersion -or $localVersion -eq "0.0.0") {
            $needDownload = $true
            Write_LogEntry -Message "Keine lokale Version vorhanden -> Download erforderlich." -Level "DEBUG"
        } else {
            $needDownload = ([version]($webVersion) -gt [version]($localVersion))
            Write_LogEntry -Message "Versionsvergleich: Online ($webVersion) > Lokal ($localVersion) ? $needDownload" -Level "DEBUG"
        }
    } catch {
        $needDownload = $true
        Write_LogEntry -Message "Fehler beim Versionsvergleich; nehme an, dass ein Download erforderlich ist: $($_)" -Level "WARNING"
    }

    if (-not $needDownload) {
        Write_LogEntry -Message "Online-Version ($webVersion) ist nicht neuer als lokale Version ($localVersion). Überspringe Download." -Level "INFO"
        Write-Host "Kein Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    } elseif (-not $directDownloadUrl) {
        Write_LogEntry -Message "Es wurde kein direkter Download-Link gefunden; Download nicht möglich." -Level "ERROR"
        Write-Host "Kein direkter Download-Link gefunden; Update übersprungen." -ForegroundColor Yellow
    } else {
        $newFileName = if ($assetName) { $assetName } else { "windowsdesktop-runtime-$webVersion-win-x64.exe" }
        $newFilePath = Join-Path -Path $InstallationFolder -ChildPath $newFileName

        # download to temp file first to avoid overwriting and accidental deletion
        $tempFile = "$newFilePath.part"
        Write_LogEntry -Message "Starte Herunterladen des Installers (temp): $($directDownloadUrl) -> $($tempFile)" -Level "INFO"

        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add("user-agent", $headers['User-Agent'])
            if ($headers.ContainsKey('Authorization')) { $wc.Headers.Add("Authorization", $headers['Authorization']) }
            [void](Invoke-DownloadFile -Url $directDownloadUrl -OutFile $tempFile)
            $wc.Dispose()
            Write_LogEntry -Message "Temp-Download abgeschlossen: $($tempFile)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Download $($directDownloadUrl): $($_)" -Level "ERROR"
            try { if ($wc) { $wc.Dispose() } } catch {}
            $tempFile = $null
        }

        # move temp to final location (atomic-ish) and ensure we don't delete the final file
        if ($tempFile -and (Test-Path -Path $tempFile)) {
            try {
                # Move-Item -Force will replace existing target if present
                Move-Item -Path $tempFile -Destination $newFilePath -Force -ErrorAction Stop
                Write_LogEntry -Message "Temp-Datei verschoben nach finalem Pfad: $($newFilePath)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Verschieben der Temp-Datei nach finalem Pfad: $($_)" -Level "ERROR"
            }

            # If there was a different previous local file (different filename), remove it now
            if ($localFilePath -and (Test-Path -Path $localFilePath) -and ($localFilePath -ne $newFilePath)) {
                try {
                    Remove-Item -Path $localFilePath -Force -ErrorAction Stop
                    Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFilePath)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($localFilePath): $($_)" -Level "WARNING"
                }
            }

            Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) Update erfolgreich auf Version $($webVersion)" -Level "SUCCESS"
            # update localFilePath and localVersion to reflect new file
            $localFilePath = $newFilePath
            try {
                $fileVersionInfo = (Get-Item -LiteralPath $localFilePath -ErrorAction Stop).VersionInfo
                $localVersion = ($fileVersionInfo.FileVersion -split "\." | Select-Object -First 3) -join "."
                Write_LogEntry -Message "Nach Update ermittelte lokale Dateiversion: $($localVersion) aus Datei $($localFilePath)" -Level "INFO"
            } catch {
                Write_LogEntry -Message "Nach Update: Fehler beim Ermitteln der lokalen Dateiversion: $($_)" -Level "WARNING"
            }
        } else {
            Write_LogEntry -Message "Temp-Download nicht vorhanden; Download fehlgeschlagen: $($tempFile)" -Level "ERROR"
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "red"
        }
    }

} else {
    Write_LogEntry -Message "Kein passender Online-Release gefunden; keine Aktion." -Level "WARNING"
}

Write-Host ""
Write_LogEntry -Message "Beginne Überprüfung der Registry-Installationspfade" -Level "DEBUG"
$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registry-Pfad existiert: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht gefunden: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path -and $Path.Count -gt 0) {
    # Extract version numbers from DisplayName where possible
    $installedVersions = @()
    foreach ($entry in $Path) {
        $dn = $entry.DisplayName
        if ($dn -match '([\d]+\.[\d]+(\.[\d]+)?)') {
            $installedVersions += $matches[1]
        } elseif ($entry.DisplayVersion) {
            $installedVersions += $entry.DisplayVersion
        }
    }

    if ($installedVersions.Count -gt 0) {
        # pick highest
        $installedVersion = ($installedVersions | Sort-Object {[version]($_ -replace '[^\d\.]','')} -Descending)[0]
    } else {
        $installedVersion = $null
    }

    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Gefundene installierte Version: $($installedVersion); Installationsdatei Version: $($localVersion)" -Level "INFO"

    try {
        if ($installedVersion -and $localVersion -and ([version]($installedVersion -replace '[^\d\.]','') -lt [version]($localVersion -replace '[^\d\.]',''))) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
            $Install = $true
            Write_LogEntry -Message "Install erforderlich: Install=true (installed $($installedVersion) < local $($localVersion))" -Level "INFO"
        } elseif ($installedVersion -and $localVersion -and ([version]($installedVersion -replace '[^\d\.]','') -eq [version]($localVersion -replace '[^\d\.]',''))) {
            Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
            $Install = $false
            Write_LogEntry -Message "Install nicht erforderlich: installiert gleich lokal ($($localVersion))" -Level "INFO"
        } else {
            $Install = $false
            Write_LogEntry -Message "Install nicht erforderlich: installiert neuer als lokal oder keine Versionsangaben vorhanden." -Level "WARNING"
        }
    } catch {
        $Install = $false
        Write_LogEntry -Message "Fehler beim Vergleichen der Registry-Versionsangaben: $($_). Install=false" -Level "WARNING"
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
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\WindowsDesktopRuntimesInstall.ps1" `
		-InstallationFlag
	Write_LogEntry -Message "Externer Aufruf abgeschlossen: WindowsDesktopRuntimesInstall.ps1 (mit -InstallationFlag)" -Level "DEBUG"
} elseif($Install -eq $true){
	Write_LogEntry -Message "Starte externes Installationsskript (Install=true)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\WindowsDesktopRuntimesInstall.ps1"
	Write_LogEntry -Message "Externer Aufruf abgeschlossen: WindowsDesktopRuntimesInstall.ps1" -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht" -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
