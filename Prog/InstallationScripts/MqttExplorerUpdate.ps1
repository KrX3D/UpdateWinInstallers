param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "MQTT Explorer"
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

# GitHub repo info
$GitHubUser = "thomasnordquist"
$GitHubRepo = "MQTT-Explorer"
Write_LogEntry -Message "GitHubUser: $($GitHubUser); GitHubRepo: $($GitHubRepo)" -Level "DEBUG"

# local installer pattern and detection
$installerFilter = "MQTT-Explorer*.exe"
Write_LogEntry -Message "Suche lokale Installer mit Filter: $($installerFilter) in $($InstallationFolder)" -Level "DEBUG"

$FoundFile = Get-ChildItem -Path $InstallationFolder -Filter $installerFilter -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1

if ($FoundFile) {
    Write_LogEntry -Message "Gefundene lokale Datei: $($FoundFile.FullName)" -Level "DEBUG"
    $InstallationFileName = $FoundFile.Name
    $localInstaller = $FoundFile.FullName
} else {
    Write_LogEntry -Message "Keine lokale Datei gefunden mit Pattern: $($installerFilter) in $($InstallationFolder)" -Level "WARNING"
    $InstallationFileName = $null
    $localInstaller = $null
}

# helper: normalize version string (remove leading v, -beta etc.)
function Normalize-VersionString {
    param([string]$v)
    if (-not $v) { return $null }
    # Remove leading 'v' or 'V'
    $v = $v -replace '^[vV]', ''
    # replace -beta or _beta with .
    $v = $v -replace '-beta.*$', ''
    $v = $v -replace '_beta.*$', ''
    # Trim whitespace
    $v = $v.Trim()
    return $v
}

# helper: try convert to [version]
function Try-ConvertToVersion {
    param([string]$v)
    if (-not $v) { return $null }
    $vNorm = Normalize-VersionString -v $v
    try {
        return [version]$vNorm
    } catch {
        # attempt to keep only numeric/dots
        $num = ($vNorm -replace '[^0-9\.]', '')
        if ($num) {
            try { return [version]$num } catch { return $null }
        }
        return $null
    }
}

# Get local version from installer file (if exists)
function Get-LocalInstallerVersion {
    param([string]$path)
    if (-not $path) { return $null }
    if (-not (Test-Path -Path $path)) { return $null }
    try {
        $vi = (Get-Item $path).VersionInfo
        $ver = $vi.ProductVersion
        if (-not $ver) { $ver = $vi.FileVersion }
        $ver = Normalize-VersionString -v $ver
        Write_LogEntry -Message "Lokale Installer-Version aus $path : $($ver)" -Level "DEBUG"
        return $ver
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Version aus $path : $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

# Get installed version from registry (if installed)
function Get-InstalledVersionFromRegistry {
    $RegistryPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($RegPath in $RegistryPaths) {
        if (Test-Path $RegPath) {
            try {
                $items = Get-ChildItem $RegPath -ErrorAction SilentlyContinue | Get-ItemProperty -ErrorAction SilentlyContinue
                foreach ($it in $items) {
                    if ($it.DisplayName -and $it.DisplayName -like "$ProgramName*") {
                        $dv = $it.DisplayVersion
                        if ($dv) {
                            return (Normalize-VersionString -v $dv)
                        }
                    }
                }
            } catch {
                # continue
            }
        }
    }
    return $null
}

# prepare GitHub API headers
$apiUrl = "https://api.github.com/repos/$GitHubUser/$GitHubRepo/releases/latest"
$apiHeaders = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $apiHeaders['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GithubToken übergeben - wird für API-Requests verwendet." -Level "DEBUG"
}

# function: fetch latest release (with error handling)
function Get-LatestRelease {
    param([string]$url, [hashtable]$headers)
    try {
        Write_LogEntry -Message "Rufe GitHub Releases API ab: $url" -Level "DEBUG"
        $lr = Invoke-RestMethod -Uri $url -Headers $headers -ErrorAction Stop
        return $lr
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen der GitHub Releases API: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# function: choose best asset (prefer Setup .exe, avoid zips/portable)
function Choose-BestAsset {
    param($assets)
    if (-not $assets) { return $null }

    # Candidate ordering: Setup installer executables, then any exe not archive
    # Avoid names containing 'portable' or '.zip', '.tar', '.gz'
    $exeAssets = $assets | Where-Object { $_.name -match '\.exe$' -and ($_.name -notmatch '(?i)(portable|zip|tar|\.7z|\.gz)') }

    # prefer those with 'setup' or 'Setup' or 'Installer' in name
    $preferred = $exeAssets | Where-Object { $_.name -match '(?i)(setup|installer)' } | Sort-Object { $_.name.Length } | Select-Object -First 1
    if ($preferred) { return $preferred }

    # fallback: exe that contains 'MQTT' or 'MQTT-Explorer'
    $preferred2 = $exeAssets | Where-Object { $_.name -match '(?i)mqtt' } | Sort-Object { $_.name.Length } | Select-Object -First 1
    if ($preferred2) { return $preferred2 }

    # final fallback: any exe
    $fallback = $exeAssets | Select-Object -First 1
    if ($fallback) { return $fallback }

    return $null
}

# MAIN CHECK & UPDATE
function CheckAndUpdateMQTTExplorer {
    Write_LogEntry -Message "CheckAndUpdateMQTTExplorer gestartet." -Level "INFO"

    $localVersion = Get-LocalInstallerVersion -path $localInstaller
    if (-not $localVersion) {
        Write_LogEntry -Message "Keine lokale Installer-Version gefunden; setze lokal auf null." -Level "DEBUG"
    }

    # fetch release
    $latestRelease = Get-LatestRelease -url $apiUrl -headers $apiHeaders
    if (-not $latestRelease) {
        Write_LogEntry -Message "Keine Release-Informationen verfügbar; Abbruch Update-Check." -Level "ERROR"
        return
    }

    # get online version
    $onlineTag = $latestRelease.tag_name
    $onlineVersion = Normalize-VersionString -v $onlineTag
    Write_LogEntry -Message "Online release tag: $($onlineTag); normalized: $($onlineVersion)" -Level "DEBUG"

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
    Write-Host ""

    # compare using [version] if possible
    $needUpdate = $false
    try {
        $vLocal = Try-ConvertToVersion -v $localVersion
        $vOnline = Try-ConvertToVersion -v $onlineVersion
        if ($vOnline -and $vLocal) {
            if ($vOnline -gt $vLocal) { $needUpdate = $true }
        } elseif ($vOnline -and -not $vLocal) {
            # no local version -> update needed
            $needUpdate = $true
        } else {
            # fallback to string comparison
            if ($onlineVersion -ne $localVersion) { $needUpdate = $true }
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Vergleichen der Versionen: $($_.Exception.Message)" -Level "WARNING"
        if ($onlineVersion -ne $localVersion) { $needUpdate = $true }
    }

    if (-not $needUpdate) {
        Write_LogEntry -Message "Kein Update nötig (Online: $onlineVersion; Lokal: $localVersion)." -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        return
    }

    Write_LogEntry -Message "Update erforderlich: Online $onlineVersion > Lokal $localVersion" -Level "INFO"

    # Choose best installer asset
    $chosenAsset = Choose-BestAsset -assets $latestRelease.assets
    if (-not $chosenAsset) {
        Write_LogEntry -Message "Kein geeignetes Asset gefunden im Release." -Level "ERROR"
        return
    }

    $setupUrl = $chosenAsset.browser_download_url
    $setupName = $chosenAsset.name
    Write_LogEntry -Message "Gewähltes Asset: $setupName -> $setupUrl" -Level "DEBUG"

    # prepare download path
    $downloadPath = Join-Path -Path $InstallationFolder -ChildPath $setupName
    Write_LogEntry -Message "DownloadPath: $downloadPath" -Level "DEBUG"

    # download (use WebClient and set auth if provided)
    $wc = New-Object System.Net.WebClient
    try {
        $wc.Headers.Add("User-Agent", "InstallationScripts/1.0")
        if ($GithubToken) { $wc.Headers.Add("Authorization", "token $GithubToken") }

        Write_LogEntry -Message "Starte Download: $setupUrl -> $downloadPath" -Level "INFO"
        [void](Invoke-DownloadFile -Url $setupUrl -OutFile $downloadPath)
        Write_LogEntry -Message "Download abgeschlossen: $downloadPath" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Herunterladen $setupUrl : $($_.Exception.Message)" -Level "ERROR"
        return
    } finally {
        if ($wc) { $wc.Dispose() }
    }

    # Post-download: remove old installer (if exists) and replace variables
    if (Test-Path -Path $downloadPath) {
        try {
            if ($localInstaller -and (Test-Path -Path $localInstaller)) {
                Remove-Item -Path $localInstaller -Force -ErrorAction Stop
                Write_LogEntry -Message "Alte Installationsdatei entfernt: $localInstaller" -Level "DEBUG"
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $localInstaller : $($_.Exception.Message)" -Level "WARNING"
        }

        # update localInstaller variable for later
        $localInstaller = $downloadPath
        $InstallationFileName = Split-Path -Path $downloadPath -Leaf

        Write-Host "$ProgramName wurde aktualisiert auf $onlineVersion" -foregroundcolor "Green"
        Write_LogEntry -Message "$ProgramName Update erfolgreich: $downloadPath" -Level "SUCCESS"
    } else {
        Write_LogEntry -Message "Download-Datei wurde nicht gefunden nach dem Download: $downloadPath" -Level "ERROR"
    }
}

# run the check/update
Write_LogEntry -Message "Aufruf CheckAndUpdateMQTTExplorer" -Level "INFO"
CheckAndUpdateMQTTExplorer
Write_LogEntry -Message "Rückkehr nach CheckAndUpdateMQTTExplorer()" -Level "DEBUG"

Write-Host ""

# Re-evaluate local installer & version for install decision
if ($localInstaller -and (Test-Path -Path $localInstaller)) {
    $localVersion = Get-LocalInstallerVersion -path $localInstaller
    Write_LogEntry -Message "Ermittelte lokale Version erneut: $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein lokaler Installer gefunden nach Update-Versuch." -Level "DEBUG"
    $localVersion = $null
}

# Get installed version from registry
$installedVersion = Get-InstalledVersionFromRegistry
Write_LogEntry -Message "Gefundene installierte Version (Registry): $($installedVersion)" -Level "DEBUG"

# decide whether to install
$Install = $false
if ($installedVersion -and $localVersion) {
    try {
        if ([version](Try-ConvertToVersion -v $installedVersion) -lt [version](Try-ConvertToVersion -v $localVersion)) {
            $Install = $true
        }
    } catch {
        if ($installedVersion -ne $localVersion) { $Install = $true }
    }
} elseif (-not $installedVersion -and $localVersion) {
    # not installed but installer exists => install when flag or desired
    $Install = $false
}

Write_LogEntry -Message "Installationsentscheidung: Install = $($Install); InstallationFlag = $($InstallationFlag)" -Level "DEBUG"

# Install if needed / if flag set
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript mit -InstallationFlag" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\MqttExplorerInstallation.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen." -Level "DEBUG"
} elseif ($Install) {
    Write_LogEntry -Message "Starte externes Installations-Skript (Update)" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\MqttExplorerInstallation.ps1"
    Write_LogEntry -Message "Externes Installations-Skript aufgerufen (Update)." -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
