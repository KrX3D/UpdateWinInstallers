param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "ImageGlass"
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

$localFileWildcard = "ImageGlass*.msi"

$githubApiUrl = "https://api.github.com/repos/d2phap/ImageGlass/releases/latest"
Write_LogEntry -Message "Lokaler Dateiwildcard: $($localFileWildcard); GitHub API URL: $($githubApiUrl)" -Level "DEBUG"

# Get the local file matching the wildcard pattern
$localFilePath = $null
try {
    if (Test-Path -Path $InstallationFolder) {
        $localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard -File -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    } else {
        Write_LogEntry -Message "Installationsordner existiert nicht: $($InstallationFolder)" -Level "WARNING"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Auslesen lokaler Dateien: $($_)" -Level "WARNING"
}

$localFilePathStr = if ($localFilePath) { $localFilePath } else { "<none>" }
Write_LogEntry -Message "Gefundener lokaler Datei-Pfad: $localFilePathStr" -Level "DEBUG"

if (-not $localFilePath) {
    Write_LogEntry -Message "Keine lokale Datei gefunden, Wildcard: $($localFileWildcard) im Ordner: $($InstallationFolder)" -Level "WARNING"
    # localVersion bleibt null/0.0.0 in that case
    $localVersion = "0.0.0"
} else {
    # Extract local file version from the filename (pattern _Major.Minor.Build.Revision)
    $localVersion = [regex]::Match($localFilePath, '(?<=_)\d+\.\d+\.\d+\.\d+').Value
    if (-not $localVersion) {
        # fallback to reading file VersionInfo when possible
        try {
            $localVersion = (Get-Item -LiteralPath $localFilePath -ErrorAction Stop).VersionInfo.ProductVersion
        } catch {
            $localVersion = "0.0.0"
        }
    }
    Write_LogEntry -Message "Lokale Version extrahiert aus Datei $($localFilePath): $($localVersion)" -Level "DEBUG"
}

# Prepare GitHub API headers, use token if configured
$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein GitHub Token gefunden; benutze anonyme API-Anfragen (niedrigere Rate-Limits)." -Level "DEBUG"
}

# Get the latest release information from GitHub
Write_LogEntry -Message "Rufe GitHub Release-Info ab: $($githubApiUrl)" -Level "INFO"
$releaseInfo = $null
try {
    $releaseInfo = Invoke-RestMethod -Uri $githubApiUrl -Headers $headers -ErrorAction Stop
    $hasReleaseInfo = [bool]$releaseInfo
    Write_LogEntry -Message "GitHub Response empfangen; Existiert ReleaseInfo: $hasReleaseInfo" -Level "DEBUG"
} catch {
    $errMsg = $_.Exception.Message
    # try to extract a useful body if available
    try {
        if ($_.Exception.Response) {
            $rs = $_.Exception.Response.GetResponseStream()
            $rr = [System.IO.StreamReader]::new($rs)
            $rb = $rr.ReadToEnd(); $rr.Close(); $rs.Close()
            try { $rbj = $rb | ConvertFrom-Json -ErrorAction Stop; if ($rbj.message) { $errMsg = $rbj.message } } catch {}
        }
    } catch {}
    Write_LogEntry -Message "Fehler beim Abruf der GitHub-Release-Info: $($errMsg)" -Level "ERROR"
    $releaseInfo = $null
}

if ($releaseInfo) {
    # normalize online version (strip leading v if present)
    $onlineVersionRaw = $releaseInfo.tag_name
    $onlineVersion = ($onlineVersionRaw -replace '^v','')
    Write_LogEntry -Message "Online Version (Tag): $($onlineVersionRaw) -> normalisiert: $($onlineVersion); Lokale Version: $($localVersion)" -Level "INFO"

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
    Write-Host ""

    # compare versions (use [version] where possible)
    $isSame = $false
    $needDownload = $false
    try {
        if ([version]$onlineVersion -gt [version]$localVersion) { $needDownload = $true }
        elseif ([version]$onlineVersion -eq [version]$localVersion) { $isSame = $true }
    } catch {
        # if parsing fails, fallback to string comparison
        if ($onlineVersion -ne $localVersion) { $needDownload = $true } else { $isSame = $true }
    }

    if ($isSame) {
        Write_LogEntry -Message "Keine neuere Version gefunden. Online $($onlineVersion) == Lokal $($localVersion)" -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    } elseif (-not $needDownload) {
        Write_LogEntry -Message "Versionsvergleich ergab: kein Download erforderlich (unsicherer Vergleich)." -Level "WARNING"
    } else {
        # Find the asset with the download URL and exclude versions with "delete" in the filename
        $asset = $releaseInfo.assets | Where-Object { $_.name -notlike "*delete*"  -and $_.name -like "*x64*" -and $_.name -like "*.msi*" } | Select-Object -First 1
        $hasAsset = [bool]$asset
        Write_LogEntry -Message "Suche passende Asset in Release; Gefunden: $hasAsset" -Level "DEBUG"

        if ($asset) {
            $downloadUrl = $asset.browser_download_url
            $downloadFileName = $asset.name
            Write_LogEntry -Message "Ermitteltes Asset: $($downloadFileName); Download-URL: $($downloadUrl)" -Level "INFO"

            # safe download: write to temp .part then Move-Item
            $downloadPath = Join-Path $InstallationFolder $downloadFileName
            $tempPath = "$downloadPath.part"

            Write_LogEntry -Message "Starte Download (temp): $($downloadUrl) -> $($tempPath)" -Level "INFO"
            $downloadSucceeded = $false
            try {
                $wc = New-Object System.Net.WebClient
                # pass headers if needed (User-Agent already included above)
                if ($headers.ContainsKey('User-Agent')) { $wc.Headers.Add('user-agent', $headers['User-Agent']) }
                if ($headers.ContainsKey('Authorization')) { $wc.Headers.Add('Authorization', $headers['Authorization']) }
                [void](Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath)
                $wc.Dispose()
                Write_LogEntry -Message "Temp-Download abgeschlossen: $($tempPath)" -Level "DEBUG"
                $downloadSucceeded = Test-Path -Path $tempPath
            } catch {
                Write_LogEntry -Message "Fehler beim Herunterladen $($downloadUrl): $($_)" -Level "ERROR"
                try { if ($wc) { $wc.Dispose() } } catch {}
                $downloadSucceeded = $false
            }

            if ($downloadSucceeded) {
                try {
                    # move to final path atomically-ish
                    Move-Item -Path $tempPath -Destination $downloadPath -Force -ErrorAction Stop
                    Write_LogEntry -Message "Temp-Datei verschoben nach finalem Pfad: $($downloadPath)" -Level "DEBUG"

                    # If there was a different previous local file (different filename), remove it now
                    if ($localFilePath -and (Test-Path -Path $localFilePath) -and ($localFilePath -ne $downloadPath)) {
                        try {
                            Remove-Item -Path $localFilePath -Force -ErrorAction Stop
                            Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFilePath)" -Level "DEBUG"
                        } catch {
                            Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($localFilePath): $($_)" -Level "WARNING"
                        }
                    }

                    Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                    Write_LogEntry -Message "$($ProgramName) Update abgeschlossen; neue Datei: $($downloadPath)" -Level "SUCCESS"

                    # update local vars
                    $localFilePath = $downloadPath
                    try {
                        $localVersion = [regex]::Match($localFilePath, '(?<=_)\d+\.\d+\.\d+\.\d+').Value
                        if (-not $localVersion) { $localVersion = (Get-Item -LiteralPath $localFilePath).VersionInfo.ProductVersion }
                        Write_LogEntry -Message "Nach Update: Lokale Version ermittelt: $($localVersion)" -Level "DEBUG"
                    } catch {
                        Write_LogEntry -Message "Nach Update: Fehler beim Ermitteln der lokalen Version: $($_)" -Level "WARNING"
                    }
                } catch {
                    Write_LogEntry -Message "Fehler beim Verschieben/Abschließen des Downloads: $($_)" -Level "ERROR"
                    Write-Host "Download konnte nicht finalisiert werden." -ForegroundColor Red
                }
            } else {
                Write_LogEntry -Message "Temp-Download nicht vorhanden; Download fehlgeschlagen: $($tempPath)" -Level "ERROR"
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "red"
                # cleanup partial file if exists
                try { if (Test-Path $tempPath) { Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue } } catch {}
            }
        } else {
            Write_LogEntry -Message "Kein passendes Asset im Release gefunden (Filter: not *delete*, *x64*, *.msi)." -Level "WARNING"
            Write-Host "Kein passendes Release-Asset vorhanden; Update übersprungen." -ForegroundColor Yellow
        }
    }
} else {
    Write_LogEntry -Message "Keine Release-Informationen von GitHub erhalten für URL: $($githubApiUrl)" -Level "ERROR"
    Write-Host "Fehler beim Abrufen der Release-Informationen; Update übersprungen." -ForegroundColor Yellow
}

Write-Host ""
Write_LogEntry -Message "Erneute Bestimmung lokaler Datei für Installationsprüfung (Wildcard: $($localFileWildcard))." -Level "DEBUG"

#Check Installed Version / Install if neded
$FoundFile = $null
try {
    if (Test-Path -Path $InstallationFolder) {
        $FoundFile = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard -File | Select-Object -First 1 -ExpandProperty FullName
    }
} catch {
    Write_LogEntry -Message "Fehler beim Auffinden der lokalen Datei für Installationsprüfung: $($_)" -Level "WARNING"
}

$foundFileStr = if ($FoundFile) { $FoundFile } else { "<none>" }
Write_LogEntry -Message "Gefundene Datei für Installationsprüfung: $foundFileStr" -Level "DEBUG"

if ($FoundFile) {
    $localVersion = [regex]::Match($FoundFile, '(?<=_)\d+\.\d+\.\d+\.\d+').Value
    if (-not $localVersion) {
        try { $localVersion = (Get-Item -LiteralPath $FoundFile).VersionInfo.ProductVersion } catch { $localVersion = "0.0.0" }
    }
    Write_LogEntry -Message "Lokale Version (für Installationsprüfung) extrahiert: $($localVersion)" -Level "DEBUG"
} else {
    $localVersion = "0.0.0"
    Write_LogEntry -Message "Keine lokale Datei für Installationsprüfung gefunden; setze lokale Version auf $($localVersion)" -Level "DEBUG"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Zu prüfende Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write_LogEntry -Message "Gefundene installierte Version in Registry: $($installedVersion)" -Level "INFO"
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"

    if ([version]$installedVersion -lt [version]$localVersion) {
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist älter als lokale Datei ($($localVersion)). Markiere Installation." -Level "INFO"
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
        $Install = $true
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write_LogEntry -Message "Installierte Version entspricht lokaler Datei: $($installedVersion)" -Level "DEBUG"
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
        $Install = $false
    } else {
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Datei ($($localVersion)). Kein Update nötig." -Level "WARNING"
        $Install = $false
    }
} else {
    Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden. Install-Flag auf $($false) gesetzt." -Level "INFO"
    $Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ImageGlassInstall.ps1 mit Parameter -InstallationFlag" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\ImageGlassInstall.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ImageGlassInstall.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Install Flag true: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ImageGlassInstall.ps1" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\ImageGlassInstall.ps1"
    Write_LogEntry -Message "Externes Installations-Skript für ImageGlass aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\ImageGlassInstall.ps1" -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
