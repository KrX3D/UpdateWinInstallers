param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Git"
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
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PSScriptRoot: $($PSScriptRoot)" -Level "DEBUG"

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
    Write_LogEntry -Message "Konfigurationsdatei gefunden und geladen: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit 1
}

$localInstallerPath = "$InstallationFolder\Git-*-64-bit.exe"
Write_LogEntry -Message "Suche lokale Installer unter: $($localInstallerPath)" -Level "DEBUG"

# Get the local installer file (most recent)
$localInstaller = $null
try {
    $localInstaller = Get-ChildItem -Path $localInstallerPath -File -ErrorAction SilentlyContinue | Select-Object -First 1
} catch {
    Write_LogEntry -Message "Fehler beim Suchen lokaler Installer: $($_)" -Level "WARNING"
}

if ($localInstaller -ne $null) {
    Write_LogEntry -Message "Gefundene lokale Installer-Datei: $($localInstaller.FullName)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Gefundene lokale Installer-Datei: <none>" -Level "DEBUG"
}

# Determine local version: try file product version first, then fallback to filename parse, else 0.0.0
$localVersion = "0.0.0"
if ($localInstaller) {
    try {
        $prop = (Get-ItemProperty -Path $localInstaller.FullName -ErrorAction Stop).VersionInfo
        if ($prop.ProductVersion) {
            # take only first three components for robust comparison
            $pv = ($prop.ProductVersion -split '\.')[0..2] -join '.'
            $localVersion = $pv
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Produktversion vom lokalen Installer: $($_)" -Level "DEBUG"
    }

    if ($localVersion -eq "0.0.0") {
        # fallback: try to parse version from filename like Git-2.40.1-64-bit.exe
        $fn = $localInstaller.Name
        $m = [regex]::Match($fn, 'Git-(\d+\.\d+\.\d+(?:\.\d+)?)')
        if ($m.Success) {
            # keep first three components
            $v = ($m.Groups[1].Value -split '\.')[0..2] -join '.'
            $localVersion = $v
        }
    }
}

Write_LogEntry -Message "Lokale Installer-Version ermittelt: $($localVersion)" -Level "DEBUG"

# Retrieve the latest Git version from GitHub releases (use token header if available)
$releasesURL = "https://api.github.com/repos/git-for-windows/git/releases/latest"
Write_LogEntry -Message "Rufe GitHub Releases API ab: $($releasesURL)" -Level "INFO"

$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein GitHub Token gefunden; benutze anonyme API-Anfragen (niedrigere Rate-Limits)." -Level "DEBUG"
}

$latestRelease = $null
try {
    $latestRelease = Invoke-RestMethod -Uri $releasesURL -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub Releases API Aufruf abgeschlossen. Verarbeite Assets." -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abruf der GitHub Releases API: $($_)" -Level "ERROR"
    $latestRelease = $null
}

# Extract latest version from assets
$latestVersion = $null
$selectedAsset = $null

if ($latestRelease) {
    try {
        # pattern: Git-<version>-64-bit.exe
        $versionPattern = 'Git-(\d+\.\d+\.\d+(?:\.\d+)?)\-64-bit\.exe'
        if ($latestRelease.assets -and $latestRelease.assets.Count -gt 0) {
            # find assets matching the 64-bit installer pattern
            $candidates = $latestRelease.assets | Where-Object { $_.name -match $versionPattern }
            if ($candidates.Count -gt 0) {
                # choose highest version by parsing numeric parts
                $assetList = $candidates | ForEach-Object {
                    $m = [regex]::Match($_.name, $versionPattern)
                    [PSCustomObject]@{
                        Asset = $_
                        VerString = $m.Groups[1].Value
                        VerObj = try { [version](($m.Groups[1].Value -split '\.')[0..2] -join '.') } catch { [version]'0.0.0' }
                    }
                }
                $assetList = $assetList | Sort-Object -Property VerObj -Descending
                $selectedAsset = $assetList[0].Asset
                $latestVersion = $assetList[0].VerString
                # normalize to three components
                $latestVersion = ($latestVersion -split '\.')[0..2] -join '.'
            } else {
                # fallback: try any *64-bit.exe asset
                $any = $latestRelease.assets | Where-Object { $_.name -like '*64-bit.exe' } | Select-Object -First 1
                if ($any) {
                    $selectedAsset = $any
                    $m = [regex]::Match($any.name, $versionPattern)
                    if ($m.Success) {
                        $latestVersion = ($m.Groups[1].Value -split '\.')[0..2] -join '.'
                    } else {
                        # attempt to get tag_name fallback
                        $tag = $latestRelease.tag_name -replace '^v',''
                        try { $latestVersion = ($tag -split '\.')[0..2] -join '.' } catch { $latestVersion = $tag }
                    }
                }
            }
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Extrahieren der Version/Assets: $($_)" -Level "ERROR"
    }
}

if ($selectedAsset -ne $null) {
    Write_LogEntry -Message "Ermittelte Online-Version: $($latestVersion) ; ausgewähltes Asset: $($selectedAsset.name)" -Level "INFO"
} else {
    Write_LogEntry -Message "Ermittelte Online-Version: $($latestVersion) ; ausgewähltes Asset: <none>" -Level "INFO"
}

Write-Host ""
Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
Write-Host ""

# Compare the local and latest versions (using [version] when possible)
$needDownload = $false
try {
    if ($latestVersion -and $localVersion) {
        if ([version]$latestVersion -gt [version]$localVersion) { $needDownload = $true }
    } elseif ($latestVersion -and -not $localVersion) {
        $needDownload = $true
    }
} catch {
    # fallback string compare
    if ($latestVersion -ne $localVersion) { $needDownload = $true }
}

if ($needDownload -and $selectedAsset) {
    Write_LogEntry -Message "Update verfügbar: Online $($latestVersion) > Lokal $($localVersion). Starte Download-Vorbereitung." -Level "INFO"

    $downloadURL = $selectedAsset.browser_download_url
    Write_LogEntry -Message "Ermittelter Download-URL: $($downloadURL)" -Level "DEBUG"

    $downloadedFileName = [System.IO.Path]::GetFileName($downloadURL)
    $downloadPath = Join-Path -Path $InstallationFolder -ChildPath $downloadedFileName
    $tempPath = "$downloadPath.part"

    Write_LogEntry -Message "Zielpfad für Download: $($downloadPath) (temp: $($tempPath))" -Level "DEBUG"

    $downloadOk = $false
    try {
        $wc = New-Object System.Net.WebClient
        if ($headers.ContainsKey('User-Agent')) { $wc.Headers.Add('User-Agent', $headers['User-Agent']) }
        if ($headers.ContainsKey('Authorization')) { $wc.Headers.Add('Authorization', $headers['Authorization']) }
        Write_LogEntry -Message "Starte Download von $($downloadURL) nach $($tempPath)" -Level "INFO"
        [void](Invoke-DownloadFile -Url $downloadURL -OutFile $tempPath)
        $wc.Dispose()
        $downloadOk = Test-Path -Path $tempPath
        Write_LogEntry -Message "Download abgeschlossen; temp existiert: $($downloadOk)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Herunterladen: $($_)" -Level "ERROR"
        try { if ($wc) { $wc.Dispose() } } catch {}
        $downloadOk = $false
    }

    if ($downloadOk) {
        try {
            Move-Item -Path $tempPath -Destination $downloadPath -Force -ErrorAction Stop
            Write_LogEntry -Message "Temp-Datei verschoben nach: $($downloadPath)" -Level "DEBUG"

            # Remove old installer if present and different filename
            if ($localInstaller -and (Test-Path -Path $localInstaller.FullName) -and ($localInstaller.FullName -ne $downloadPath)) {
                try {
                    Remove-Item -Path $localInstaller.FullName -Force -ErrorAction Stop
                    Write_LogEntry -Message "Alter Installer entfernt: $($localInstaller.FullName)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message "Fehler beim Entfernen alter Installer-Datei: $($_)" -Level "WARNING"
                }
            }

            Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) Update-Vorgang abgeschlossen und Dateien ersetzt. Neue Datei: $($downloadPath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Finalisieren des Downloads/Move: $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Download fehlgeschlagen: temp-Datei nicht vorhanden nach Download." -Level "ERROR"
        try { if (Test-Path $tempPath) { Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue } } catch {}
        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "red"
    }
} else {
    Write_LogEntry -Message "Kein Update erforderlich oder kein Asset gefunden. needDownload=$($needDownload); assetFound=$($selectedAsset -ne $null)" -Level "INFO"
    Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
}

Write-Host ""
Write_LogEntry -Message "Starte Überprüfung installierter Versionen in Registry." -Level "DEBUG"

#Check Installed Version / Install if needed
$FoundFile = Get-ChildItem -Path $localInstallerPath -File -ErrorAction SilentlyContinue | Select-Object -First 1
if ($FoundFile) {
    Write_LogEntry -Message "Gefundene Datei für Installation: $($FoundFile.Name)" -Level "DEBUG"
    $InstallationFileName = $FoundFile.Name
    $localInstaller = Join-Path -Path $InstallationFolder -ChildPath $InstallationFileName
    Write_LogEntry -Message "Berechneter Pfad zur Installationsdatei: $($localInstaller)" -Level "DEBUG"
    try {
        $localVersion = (Get-Item $localInstaller -ErrorAction Stop).VersionInfo.ProductVersion
    } catch {
        # fallback filename parse
        $m = [regex]::Match($InstallationFileName, 'Git-(\d+\.\d+\.\d+(?:\.\d+)?)')
        if ($m.Success) { $localVersion = ($m.Groups[1].Value -split '\.')[0..2] -join '.' } else { $localVersion = "0.0.0" }
    }
    Write_LogEntry -Message "Lokale Installationsdatei Version: $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei für Registry-Check gefunden." -Level "DEBUG"
    $localVersion = $localVersion -or "0.0.0"
}

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Zu prüfende Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Gefundene installierte Version von $($ProgramName): $($installedVersion). Installationsdatei Version: $($localVersion)" -Level "INFO"

    try {
        if ([version]$installedVersion -lt [version]$localVersion) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
            Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist älter als lokale Installationsdatei ($($localVersion)). Markiere Installation." -Level "INFO"
            $Install = $true
        } elseif ([version]$installedVersion -eq [version]$localVersion) {
            Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
            Write_LogEntry -Message "Installierte Version ist aktuell: $($installedVersion)" -Level "DEBUG"
            $Install = $false
        } else {
            Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Installationsdatei ($($localVersion)). Kein Update nötig." -Level "WARNING"
            $Install = $false
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Vergleichen der Versionsangaben in Registry: $($_). Install=false" -Level "WARNING"
        $Install = $false
    }
} else {
    Write_LogEntry -Message "$($ProgramName) nicht in der Registry gefunden. Install-Flag wird auf $($false) gesetzt." -Level "INFO"
    $Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GitInstall.ps1 mit Parameter -InstallationFlag" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\GitInstall.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GitInstall.ps1" -Level "DEBUG"
}
elseif($Install -eq $true){
    Write_LogEntry -Message "Install Flag true: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GitInstall.ps1" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\GitInstall.ps1"
    Write_LogEntry -Message "Externes Installations-Skript für Git aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GitInstall.ps1" -Level "DEBUG"
}
Write-Host ""

Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
    Finalize_LogSession | Out-Null
}
# === Ende Logger-Footer ===
