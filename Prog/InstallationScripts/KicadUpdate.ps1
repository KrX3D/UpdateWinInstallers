param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "KiCad"
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
    Write_LogEntry -Message "Konfigurationsdatei gefunden und geladen: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit 1
}

$InstallationFolder = Join-Path -Path $InstallationFolder -ChildPath "Kicad"

# GitHub repo info
$GitHubUser = "KiCad"
$GitHubRepo = "kicad-source-mirror"
Write_LogEntry -Message "GitHubUser: $($GitHubUser); GitHubRepo: $($GitHubRepo)" -Level "DEBUG"

# Local installer pattern & detection
$InstallationFilePathPattern = Join-Path -Path $InstallationFolder -ChildPath "kicad*.exe"
Write_LogEntry -Message "Suche lokale Installer mit Muster: $($InstallationFilePathPattern)" -Level "DEBUG"

$FoundFile = Get-ChildItem -Path $InstallationFolder -Filter "kicad*.exe" -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1

if ($FoundFile) {
    $InstallationFileName = $FoundFile.Name
    $localInstaller = $FoundFile.FullName
    Write_LogEntry -Message "Gefundene lokale Datei: $($localInstaller)" -Level "DEBUG"
} else {
    $InstallationFileName = $null
    $localInstaller = $null
    Write_LogEntry -Message "Keine lokale Installationsdatei gefunden mit Muster: $($InstallationFilePathPattern)" -Level "WARNING"
}

# Helper: normalize version string (remove v, -beta etc.)
function Normalize-VersionString {
    param([string]$v)
    if (-not $v) { return $null }
    $v = $v -replace '^[vV]', ''
    $v = $v -replace '-beta.*$', ''
    $v = $v -replace '_beta.*$', ''
    $v = $v.Trim()
    return $v
}

function Try-ConvertToVersion {
    param([string]$v)
    if (-not $v) { return $null }
    $vNorm = Normalize-VersionString -v $v
    try {
        return [version]$vNorm
    } catch {
        $num = ($vNorm -replace '[^0-9\.]', '')
        try { return [version]$num } catch { return $null }
    }
}

# Get local installer version safely
function Get-LocalInstallerVersion {
    param([string]$path)
    if (-not $path) { return $null }
    if (-not (Test-Path -Path $path)) { return $null }
    try {
        $vi = (Get-Item $path).VersionInfo
        $ver = $vi.ProductVersion
        if (-not $ver) { $ver = $vi.FileVersion }
        return Normalize-VersionString -v $ver
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Produktversion für $path : $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

# Get installed version from registry (if present)
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
                        if ($it.DisplayVersion) {
                            return Normalize-VersionString -v $it.DisplayVersion
                        }
                    }
                }
            } catch {
                # ignore and continue
            }
        }
    }
    return $null
}

# Build headers for GitHub API calls
$apiUrl = "https://api.github.com/repos/$GitHubUser/$GitHubRepo/releases/latest"
$apiHeaders = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $apiHeaders['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token bereitgestellt - verwende autorisierte Anfragen." -Level "DEBUG"
}

# Choose best asset from release assets
function Choose-BestAsset {
    param($assets)
    if (-not $assets) { return $null }

    # filter out archives and arm builds; look for windows 64-bit installers
    $candidates = $assets | Where-Object {
        $_.name -match '\.exe$' -and
        ($_.name -notmatch '(?i)(portable|zip|tar|\.7z|\.gz|msi$)') -and
        ($_.name -notmatch '(?i)arm64')
    }

    if (-not $candidates -or $candidates.Count -eq 0) {
        # fallback: any exe
        $candidates = $assets | Where-Object { $_.name -match '\.exe$' }
    }

    # prefer names that include 'win', 'windows', '64' or 'x64' or 'Setup'
    $preferred = $candidates | Where-Object { $_.name -match '(?i)(win|windows|x64|64|setup|installer)' } |
                 Sort-Object { $_.name.Length } | Select-Object -First 1

    if ($preferred) { return $preferred }

    # fallback to first candidate
    return $candidates | Select-Object -First 1
}

# Fetch latest release (robust)
function Get-LatestRelease {
    param([string]$url, [hashtable]$headers)
    try {
        Write_LogEntry -Message "Rufe GitHub Releases API ab: $url" -Level "DEBUG"
        $release = Invoke-RestMethod -Uri $url -Headers $headers -ErrorAction Stop
        return $release
    } catch {
        $err = $_.Exception.Message
        Write_LogEntry -Message "Fehler beim Abrufen der GitHub API: $err" -Level "ERROR"
        return $null
    }
}

# Main update check
function CheckAndUpdateKiCad {
    Write_LogEntry -Message "CheckAndUpdateKiCad gestartet." -Level "INFO"

    $LocalVersion = Get-LocalInstallerVersion -path $localInstaller
    Write_LogEntry -Message "Lokale Version ermittelt: $($LocalVersion)" -Level "DEBUG"

    $LatestRelease = Get-LatestRelease -url $apiUrl -headers $apiHeaders
    if (-not $LatestRelease) {
        Write_LogEntry -Message "Keine Release-Informationen gefunden; Abbruch." -Level "ERROR"
        return
    }

    $OnlineTag = $LatestRelease.tag_name
    $OnlineVersion = Normalize-VersionString -v $OnlineTag
    Write_LogEntry -Message "Online-Release Tag: $($OnlineTag); Normalized: $($OnlineVersion)" -Level "DEBUG"

    Write-Host ""
    Write-Host "Lokale Version: $LocalVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $OnlineVersion" -foregroundcolor "Cyan"
    Write-Host ""

    # decide whether update is required
    $needUpdate = $false
    try {
        $vLocal = Try-ConvertToVersion -v $LocalVersion
        $vOnline = Try-ConvertToVersion -v $OnlineVersion
        if ($vOnline -and $vLocal) {
            if ($vOnline -gt $vLocal) { $needUpdate = $true }
        } elseif ($vOnline -and -not $vLocal) {
            $needUpdate = $true
        } else {
            if ($OnlineVersion -ne $LocalVersion) { $needUpdate = $true }
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Vergleich der Versionen: $($_.Exception.Message)" -Level "WARNING"
        if ($OnlineVersion -ne $LocalVersion) { $needUpdate = $true }
    }

    if (-not $needUpdate) {
        Write_LogEntry -Message "Kein Update nötig (Online: $OnlineVersion; Lokal: $LocalVersion)." -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        return
    }

    Write_LogEntry -Message "Update erkannt: Online $OnlineVersion > Lokal $LocalVersion" -Level "INFO"

    # pick best asset
    $chosenAsset = Choose-BestAsset -assets $LatestRelease.assets
    if (-not $chosenAsset) {
        Write_LogEntry -Message "Kein geeignetes Asset im Release gefunden." -Level "ERROR"
        return
    }
    $InstallerURL = $chosenAsset.browser_download_url
    $FileName = $chosenAsset.name
    Write_LogEntry -Message "Gewähltes Asset: $($FileName) -> $($InstallerURL)" -Level "DEBUG"

    # prepare download path
    $DownloadPath = Join-Path -Path $InstallationFolder -ChildPath $FileName

    # download with WebClient and optional auth
    $wc = New-Object System.Net.WebClient
    try {
        $wc.Headers.Add("User-Agent", "InstallationScripts/1.0")
        if ($GithubToken) { $wc.Headers.Add("Authorization", "token $GithubToken") }
        Write_LogEntry -Message "Starte Download: $InstallerURL -> $DownloadPath" -Level "INFO"
        [void](Invoke-DownloadFile -Url $InstallerURL -OutFile $DownloadPath)
        Write_LogEntry -Message "Download abgeschlossen: $DownloadPath" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Herunterladen $InstallerURL : $($_.Exception.Message)" -Level "ERROR"
        return
    } finally {
        if ($wc) { $wc.Dispose() }
    }

    if (Test-Path -Path $DownloadPath) {
        try {
            if ($localInstaller -and (Test-Path -Path $localInstaller)) {
                Remove-Item -Path $localInstaller -Force -ErrorAction Stop
                Write_LogEntry -Message "Alte Installationsdatei entfernt: $localInstaller" -Level "DEBUG"
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Entfernen alter Installationsdatei $localInstaller : $($_.Exception.Message)" -Level "WARNING"
        }

        # set new local installer variables
        $localInstaller = $DownloadPath
        $InstallationFileName = Split-Path -Path $DownloadPath -Leaf

        Write-Host "$ProgramName wurde aktualisiert auf $OnlineVersion" -foregroundcolor "Green"
        Write_LogEntry -Message "$ProgramName Update erfolgreich: $DownloadPath" -Level "SUCCESS"
    } else {
        Write_LogEntry -Message "Download-Datei nicht gefunden nach Download: $DownloadPath" -Level "ERROR"
    }
}

# Run check/update
Write_LogEntry -Message "Aufruf CheckAndUpdateKiCad" -Level "INFO"
CheckAndUpdateKiCad
Write_LogEntry -Message "Rückkehr nach CheckAndUpdateKiCad" -Level "DEBUG"

Write-Host ""

# Re-evaluate local installer & version for install decision
if ($localInstaller -and (Test-Path -Path $localInstaller)) {
    $localVersion = Get-LocalInstallerVersion -path $localInstaller
    Write_LogEntry -Message "Ermittelte lokale Version erneut: $($localVersion)" -Level "DEBUG"
} else {
    $localVersion = $null
    Write_LogEntry -Message "Kein lokaler Installer vorhanden nach Update-Versuch." -Level "DEBUG"
}

# Read installed version from registry
$installedVersion = Get-InstalledVersionFromRegistry
Write_LogEntry -Message "Gefundene installierte Version (Registry): $($installedVersion)" -Level "DEBUG"

# Decide whether to install
$Install = $false
try {
    if ($installedVersion -and $localVersion) {
        $vi = Try-ConvertToVersion -v $installedVersion
        $vl = Try-ConvertToVersion -v $localVersion
        if ($vi -and $vl -and ($vi -lt $vl)) {
            $Install = $true
        }
    } elseif (-not $installedVersion -and $localVersion) {
        # not currently installed but we have installer - do not auto-install unless flag is set
        $Install = $false
    }
} catch {
    Write_LogEntry -Message "Fehler bei Install-Entscheidung: $($_.Exception.Message)" -Level "WARNING"
}

Write_LogEntry -Message "Installationsentscheidung: Install = $($Install); InstallationFlag = $($InstallationFlag)" -Level "DEBUG"

# Install if requested
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript mit -InstallationFlag" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\KicadInstallation.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen." -Level "DEBUG"
} elseif ($Install) {
    Write_LogEntry -Message "Starte externes Installations-Skript (Update)" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\KicadInstallation.ps1"
    Write_LogEntry -Message "Externes Installations-Skript aufgerufen (Update)." -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
    Finalize_LogSession | Out-Null
} else {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
}
# === Ende Logger-Footer ===
