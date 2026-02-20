param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Arduino"
$ScriptType = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

$InstallationFolder = "$InstallationFolder\Arduino"
Write_LogEntry -Message "InstallationFolder gesetzt auf: $($InstallationFolder)" -Level "DEBUG"

$InstallationFilePattern = "arduino*.exe"
Write_LogEntry -Message "InstallationFilePattern gesetzt auf: $($InstallationFilePattern)" -Level "DEBUG"

Write_LogEntry -Message "ProgramName gesetzt auf: $($ProgramName)" -Level "DEBUG"

# Get the local installer (exclude _old)
$localInstaller = Get-InstallerFilePath -Directory $InstallationFolder -Filter $InstallationFilePattern -ExcludeNameLike "*_old*"

if ($null -ne $localInstaller) {
    Write_LogEntry -Message "Lokaler Installer gefunden: $($localInstaller.Name)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein lokaler Installer gefunden mit Pattern $($InstallationFilePattern) in $($InstallationFolder)" -Level "WARNING"
}

# Determine local version from filename or file properties
$localFileVersion = $null
if ($localInstaller) {
    # Try product version first (executable properties)
    try {
        $productVersion = Get-InstallerFileVersion -FilePath $localInstaller.FullName -Source ProductVersion
        if ($productVersion) {
            # take first three components for consistent comparison
            $localFileVersion = (($productVersion -split '\.')[0..2] -join '.')
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Produktversion vom lokalen Installer: $($_)" -Level "DEBUG"
    }

    # If product version not available, parse from filename (robust match)
    if (-not $localFileVersion) {
        $localFileVersion = Get-InstallerFileVersion -FilePath $localInstaller.FullName -FileNameRegex '(\d+\.\d+\.\d+)' -Source FileName
    }
}

if (-not $localFileVersion) { $localFileVersion = "0.0.0" }

$localInstallerName = if ($localInstaller -ne $null) { $localInstaller.Name } else { '<none>' }
Write_LogEntry -Message "Ermittelte lokale Dateiversion aus Datei: $localInstallerName => $($localFileVersion)" -Level "DEBUG"

$githubApiUrl = "https://api.github.com/repos/arduino/arduino-ide/releases/latest"
Write_LogEntry -Message "GitHub API URL gesetzt auf: $($githubApiUrl)" -Level "DEBUG"

# Prepare headers (use Github token from PowerShellVariables.ps1 if present)
$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein GitHub Token gefunden; benutze anonyme API-Anfragen (niedrigere Rate-Limits)." -Level "DEBUG"
}

# Retrieve the latest release information from GitHub API with headers
Write_LogEntry -Message "Rufe GitHub API ab: $($githubApiUrl)" -Level "INFO"
try {
    $latestRelease = Invoke-RestMethod -Uri $githubApiUrl -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub API Antwort erhalten. Anzahl Assets: $($latestRelease.assets.Count)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abruf der GitHub API: $($_)" -Level "ERROR"
    $latestRelease = $null
}

if ($null -eq $latestRelease) {
    Write_LogEntry -Message "Keine Release-Information empfangen; Abbruch des Update-Checks." -Level "ERROR"
} else {
    # Filter the assets for Windows 64-bit exe. Use pattern matching to be tolerant.
    # Accept assets where name or browser_download_url suggests Windows + 64 and is an exe.
    $downloadAsset = $latestRelease.assets | Where-Object {
        ($_.name -match '(?i)win|windows|windows64|windows-64|windows_64|64bit|64-bit|x64') -or
        ($_.browser_download_url -match '(?i)win|windows|x64|64bit|64-bit')
    } | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1

    if ($null -ne $downloadAsset) {
        Write_LogEntry -Message "Geeignetes Download-Asset gefunden: $($downloadAsset.name) / $($downloadAsset.browser_download_url)" -Level "INFO"
    } else {
        Write_LogEntry -Message "Kein passendes Download-Asset für Windows 64bit exe gefunden in GitHub-Release." -Level "WARNING"
    }

    if ($downloadAsset) {
        # Get the download URL and filename
        $downloadURL = $downloadAsset.browser_download_url
        Write_LogEntry -Message "Download-URL gesetzt auf: $($downloadURL)" -Level "DEBUG"

        $downloadFileName = [System.IO.Path]::GetFileName($downloadURL)
        Write_LogEntry -Message "Download-FileName ermittelt: $($downloadFileName)" -Level "DEBUG"

        # Extract online version from the filename (robust match)
        $mOnline = [regex]::Match($downloadFileName, '(\d+\.\d+\.\d+)')
        if ($mOnline.Success) {
            $onlineVersion = $mOnline.Groups[1].Value
        } else {
            # fallback to tag_name if filename parse fails
            $onlineVersion = $latestRelease.tag_name -replace '^v',''
            $mTag = [regex]::Match($onlineVersion, '(\d+\.\d+\.\d+)')
            if ($mTag.Success) { $onlineVersion = $mTag.Groups[1].Value } else { $onlineVersion = $onlineVersion }
        }

        Write_LogEntry -Message "Extrahierte Online-Version aus Dateiname/Tag: $($onlineVersion)" -Level "DEBUG"

        Write-Host ""
        Write-Host "Lokale Version: $localFileVersion" -foregroundcolor "Cyan"
        Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
        Write-Host ""

        # Compare the local and online versions using [version] where possible
        $doDownload = $false
        try {
            if ([version]$onlineVersion -ne [version]$localFileVersion) {
                if ([version]$onlineVersion -gt [version]$localFileVersion) { $doDownload = $true }
            }
        } catch {
            # fallback to string compare if version cast fails
            if ($onlineVersion -ne $localFileVersion) { $doDownload = $true }
        }

        if ($doDownload) {
            Write_LogEntry -Message "Versionen unterscheiden sich. Online: $($onlineVersion) Lokal: $($localFileVersion). Starte Download." -Level "INFO"

            # Set the download path and download to a .part temporary file first
            $downloadPath = Join-Path -Path $InstallationFolder -ChildPath $downloadFileName
            $tempPath = "$downloadPath.part"
            Write_LogEntry -Message "DownloadPath gesetzt auf: $($downloadPath) (temp: $($tempPath))" -Level "DEBUG"

            $webClient = New-Object System.Net.WebClient
            if ($headers.ContainsKey('User-Agent')) { $webClient.Headers.Add('User-Agent', $headers['User-Agent']) }
            if ($headers.ContainsKey('Authorization')) { $webClient.Headers.Add('Authorization', $headers['Authorization']) }

            try {
                Write_LogEntry -Message "Starte Download von $($downloadURL) nach $($tempPath)" -Level "INFO"
                [void](Invoke-DownloadFile -Url $downloadURL -OutFile $tempPath)
                $webClient.Dispose()
                Write_LogEntry -Message "Download beendet (temp vorhanden: $([bool](Test-Path $tempPath)))" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Herunterladen: $($_)" -Level "ERROR"
                try { if ($webClient) { $webClient.Dispose() } } catch {}
            }

            # Check if the file was completely downloaded
            if (Test-Path $tempPath) {
                try {
                    Move-Item -Path $tempPath -Destination $downloadPath -Force -ErrorAction Stop
                    Write_LogEntry -Message "Downloaddatei verschoben: $($downloadPath)" -Level "DEBUG"

                    # Remove the old installer if present
                    if ($localInstaller -and (Test-Path -Path $localInstaller.FullName) -and ($localInstaller.FullName -ne $downloadPath)) {
                        try {
                            Remove-Item -Path $localInstaller.FullName -Force -ErrorAction Stop
                            Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localInstaller.FullName)" -Level "DEBUG"
                        } catch {
                            Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei: $($_)" -Level "WARNING"
                        }
                    }

                    Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                    Write_LogEntry -Message "$($ProgramName) wurde aktualisiert. Neue Datei: $($downloadPath)" -Level "SUCCESS"
                } catch {
                    Write_LogEntry -Message "Fehler beim Finalisieren des Downloads: $($_)" -Level "ERROR"
                    Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
                }
            } else {
                Write_LogEntry -Message "Download fehlgeschlagen. Temp-Datei nicht gefunden: $($tempPath)" -Level "ERROR"
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
            }
        } else {
            Write_LogEntry -Message "Kein Online Update verfügbar. Online:$($onlineVersion) Lokal:$($localFileVersion)" -Level "INFO"
            Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        }
    }
}

Write-Host ""
Write_LogEntry -Message "Abschnitt Prüfung/Download abgeschlossen." -Level "DEBUG"

#Check Installed Version / Install if needed (re-evaluate local file after potential download)
$FoundFile = Get-InstallerFilePath -Directory $InstallationFolder -Filter $InstallationFilePattern -ExcludeNameLike "*_old*"

if ($null -ne $FoundFile) {
    Write_LogEntry -Message "Gefundene Installationsdatei für Check: $($FoundFile.FullName)" -Level "DEBUG"
    $InstallationFileName = $FoundFile.Name
    $localInstallerPath = $FoundFile.FullName
    try {
        $localVersion = Get-InstallerFileVersion -FilePath $localInstallerPath -Source ProductVersion
    } catch {
        # fallback to filename parse
        $localVersion = Get-InstallerFileVersion -FilePath $localInstallerPath -FileNameRegex '(\d+\.\d+\.\d+)' -Source FileName
    }
    Write_LogEntry -Message "Lokaler Installer Pfad: $($localInstallerPath), ProduktVersion/Filename: $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine Installationsdatei für Check gefunden in $($InstallationFolder) mit Pattern $($InstallationFilePattern)" -Level "WARNING"
    $InstallationFileName = $null
    $localInstallerPath = $null
    $localVersion = $null
}

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Prüfung gesetzt: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registry-Pfad vorhanden: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht vorhanden: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"

    if ($installedVersion -and $localVersion) {
        try {
            if ([version]$installedVersion -lt [version]$localVersion) {
                Write_LogEntry -Message "Veraltete Installation erkannt. Installierte Version: $($installedVersion) ist älter als Lokale Installationsdatei Version: $($localVersion). Update wird gestartet." -Level "INFO"
                Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
                $Install = $true
            } elseif ([version]$installedVersion -eq [version]$localVersion) {
                Write_LogEntry -Message "Installierte Version entspricht lokaler Installationsdatei. Keine Aktion erforderlich." -Level "INFO"
                Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
                $Install = $false
            } else {
                Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Installationsdatei ($($localVersion)). Keine Aktion." -Level "WARNING"
                $Install = $false
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Vergleichen der Versionsangaben in Registry: $($_). Install=false" -Level "WARNING"
            $Install = $false
        }
    } else {
        Write_LogEntry -Message "Keine Versionsangaben für Vergleich vorhanden; Install=false" -Level "DEBUG"
        $Install = $false
    }
} else {
    Write_LogEntry -Message "Programm $($ProgramName) nicht in Registry gefunden." -Level "INFO"
    $Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte externes Installationsskript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1 mittels $($PSHostPath)" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Aufruf externes Skript beendet: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1" -Level "DEBUG"
}
elseif($Install -eq $true){
    Write_LogEntry -Message "Install=true. Starte externes Installationsskript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1 mittels $($PSHostPath)" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1"
    Write_LogEntry -Message "Aufruf externes Skript beendet: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1" -Level "DEBUG"
}
Write-Host ""

# ===== Logger-Footer (BEGIN) =====
Write_LogEntry -Message "Script beendet." -Level "INFO"
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
# ===== Logger-Footer (END) =====
