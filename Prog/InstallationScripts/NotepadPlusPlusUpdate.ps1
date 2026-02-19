param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Notepad++"
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
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit 1
}

$installerPattern = "npp.*.Installer.x64.exe"
$InstallationFileFile = Join-Path -Path $InstallationFolder -ChildPath $installerPattern
Write_LogEntry -Message "Suchmuster für Installer gesetzt: $($InstallationFileFile)" -Level "DEBUG"

# Try to find existing local installer (last one)
$FoundFile = Get-InstallerFilePath -Directory $InstallationFolder -Filter $installerPattern
if ($FoundFile) {
    Write_LogEntry -Message "Lokale Datei gefunden: $($FoundFile.FullName)" -Level "DEBUG"
    $InstallationFileName = $FoundFile.Name
    $installerPath = $FoundFile.FullName
} else {
    Write_LogEntry -Message "Keine lokale Datei gefunden mit Pattern: $($InstallationFileFile)" -Level "WARNING"
    $InstallationFileName = $null
    $installerPath = $null
}

# Prepare GitHub API headers (use token if provided)
$repoOwner = "notepad-plus-plus"
$repoName  = "notepad-plus-plus"
$latestReleaseApi = "https://api.github.com/repos/$repoOwner/$repoName/releases/latest"
$apiHeaders = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GithubToken) {
    $apiHeaders['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GithubToken übergeben - wird für API-Requests verwendet." -Level "DEBUG"
}

# Function: get latest release via GitHub API (preferred)
function Get-LatestReleaseFromGitHub {
    param(
        [string]$ApiUrl,
        [hashtable]$Headers
    )
    try {
        Write_LogEntry -Message "Rufe GitHub Releases API ab: $ApiUrl" -Level "DEBUG"
        $release = Invoke-RestMethod -Uri $ApiUrl -Headers $Headers -ErrorAction Stop
        return $release
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen GitHub API: $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

# Function: extract installer asset from release
function Get-InstallerAssetFromRelease {
    param(
        $releaseObject
    )
    if (-not $releaseObject) { return $null }
    $assets = $releaseObject.assets
    if (-not $assets -or $assets.Count -eq 0) {
        Write_LogEntry -Message "Keine Assets im Release gefunden." -Level "DEBUG"
        return $null
    }

    # Prefer asset matching Installer.x64.exe pattern
    $preferred = $assets | Where-Object { $_.name -match '\.Installer\.x64\.exe$' } | Sort-Object { $_.name.Length } | Select-Object -First 1
    if (-not $preferred) {
        # fallback: any x64 exe
        $preferred = $assets | Where-Object { $_.name -match 'x64.*\.exe$' -or $_.name -match '64' } | Select-Object -First 1
    }
    if (-not $preferred) {
        # ultimate fallback: first exe
        $preferred = $assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1
    }

    if ($preferred) {
        return $preferred
    } else {
        return $null
    }
}

# Function: get latest version from webpage if API fails (fallback)
function GetLatestVersionFromHtml {
    param([string]$DownloadPageUrl)
    try {
        Write_LogEntry -Message "Fallback: lese Download-Seite HTML: $DownloadPageUrl" -Level "DEBUG"
        $req = [System.Net.WebRequest]::Create($DownloadPageUrl)
        $req.Method = "GET"
        $req.UserAgent = "InstallationScripts/1.0"
        $resp = $req.GetResponse()
        $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
        $content = $sr.ReadToEnd()
        $sr.Close(); $resp.Close()
        $contentLength = $content.Length
        Write_LogEntry -Message "Download-Seite abgeholt; Länge: $contentLength" -Level "DEBUG"

        # Extract the Current Version (page text includes 'Current Version X.Y.Z')
        $versionRegex = '(?<=Current Version )[\d.]+'
        $m = [regex]::Match($content, $versionRegex)
        if ($m.Success) { return $m.Value } else { return $null }
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Download-Seite: $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

# Function: get local version from installer path
function GetLocalVersionNumber {
    param([string]$Path)
    if (-not $Path) { return $null }
    if (-not (Test-Path -Path $Path)) { return $null }
    try {
        $vi = (Get-Item $Path).VersionInfo
        $lv = $vi.FileVersion
        if (-not $lv) { return $null }
        # shorten to three parts if possible
        $parts = $lv -split '\.' | Select-Object -First 3
        if ($parts[-1] -eq '0') { $parts = $parts[0..($parts.Count - 2)] }
        return ($parts -join '.')
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Version aus $Path : $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

# Main: get latest release via API
$release = Get-LatestReleaseFromGitHub -ApiUrl $latestReleaseApi -Headers $apiHeaders

$latestVersion = $null
$chosenAsset = $null

if ($release) {
    # get tag_name
    $tagName = $release.tag_name
    # extract numeric version from tag if possible
    $vm = [regex]::Match($tagName, '\d+(\.\d+)+')
    if ($vm.Success) {
        $latestVersion = $vm.Value
    } else {
        $latestVersion = $tagName.Trim()
    }
    Write_LogEntry -Message "GitHub API: Release Tag: $($tagName); Version extracted: $($latestVersion)" -Level "DEBUG"

    $chosenAsset = Get-InstallerAssetFromRelease -releaseObject $release
    if ($chosenAsset) {
        Write_LogEntry -Message "Ausgewähltes Asset aus Release: $($chosenAsset.name) -> $($chosenAsset.browser_download_url)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Kein passendes Asset im Release gefunden." -Level "WARNING"
    }
} else {
    # fallback: scrape page
    $downloadPageUrl = "https://notepad-plus-plus.org/downloads/"
    $latestVersion = GetLatestVersionFromHtml -DownloadPageUrl $downloadPageUrl
    Write_LogEntry -Message "Fallback HTML-Version: $($latestVersion)" -Level "DEBUG"
    # if we have latestVersion, construct asset URL pattern (github path)
    if ($latestVersion) {
        $assetName = "npp.$latestVersion.Installer.x64.exe"
        $constructedUrl = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v$latestVersion/$assetName"
        $fallbackMessage = "Constructed asset URL (fallback): $constructedUrl"
        Write_LogEntry -Message $fallbackMessage -Level "DEBUG"
        $chosenAsset = @{ name = $assetName; browser_download_url = $constructedUrl }
    }
}

# Get local version
$localVersion = GetLocalVersionNumber -Path $installerPath
Write_LogEntry -Message "Lokale Version ermittelt: $($localVersion)" -Level "DEBUG"

Write-Host ""
Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
Write-Host ""

# Compare versions safely
$needDownload = $false
if ($latestVersion) {
    try {
        if ($localVersion) {
            if ([version]$latestVersion -gt [version]$localVersion) { $needDownload = $true }
        } else {
            # no local file: consider download needed
            $needDownload = $true
        }
    } catch {
        # fallback to string comparison if version parse fails
        if ($latestVersion -ne $localVersion) { $needDownload = $true }
    }
}

if ($needDownload -and $chosenAsset) {
    # prepare download path
    if ($installerPath) { $targetDir = Split-Path -Path $installerPath -Parent } else { $targetDir = $InstallationFolder }
    if (-not (Test-Path $targetDir)) { New-Item -Path $targetDir -ItemType Directory -Force | Out-Null }

    $downloadFileName = $chosenAsset.name
    if (-not $downloadFileName) {
        # if chosenAsset object was constructed as hashtable fallback
        $downloadFileName = "npp.$latestVersion.Installer.x64.exe"
    }
    $downloadUrl = $chosenAsset.browser_download_url
    $downloadPath = Join-Path -Path $targetDir -ChildPath $downloadFileName

    Write_LogEntry -Message "Starte Download: $downloadUrl -> $downloadPath" -Level "INFO"

    # Use WebClient and set Authorization header if token present (some private assets might need it)
    $wc = New-Object System.Net.WebClient
    try {
        $wc.Headers.Add("User-Agent", "InstallationScripts/1.0")
        if ($GithubToken) {
            $wc.Headers.Add("Authorization", "token $GithubToken")
        }
        [void](Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath)
        Write_LogEntry -Message "Download abgeschlossen: $downloadPath" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Download $downloadUrl : $($_.Exception.Message)" -Level "ERROR"
        $downloadPath = $null
    } finally {
        if ($wc) { $wc.Dispose() }
    }

    # handle downloaded file
    if ($downloadPath -and (Test-Path $downloadPath)) {
        try {
            if ($installerPath -and (Test-Path $installerPath)) {
                Remove-Item -Path $installerPath -Force -ErrorAction Stop
                Write_LogEntry -Message "Alte Installationsdatei entfernt: $installerPath" -Level "DEBUG"
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Entfernen alter Datei $installerPath : $($_.Exception.Message)" -Level "WARNING"
        }

        Write-Host "$ProgramName wurde aktualisiert auf $latestVersion" -foregroundcolor "Green"
        Write_LogEntry -Message "$ProgramName Update erfolgreich: $downloadFileName" -Level "SUCCESS"
        # update installerPath/InstallationFileName for further steps
        $installerPath = $downloadPath
        $InstallationFileName = Split-Path -Path $downloadPath -Leaf
    } else {
        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "Red"
        Write_LogEntry -Message "Download fehlgeschlagen oder Datei nicht vorhanden nach Download." -Level "ERROR"
    }
} elseif ($needDownload -and -not $chosenAsset) {
    Write_LogEntry -Message "Update verfügbar aber kein passendes Asset zur Verfügung." -Level "WARNING"
    Write-Host "Kein geeignetes Asset gefunden - Abbruch." -ForegroundColor "Yellow"
} else {
    Write_LogEntry -Message "Kein Update notwendig (Online: $latestVersion, Lokal: $localVersion)." -Level "INFO"
    Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
}

Write-Host ""

# Check installed version via registry and decide install
$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', `
                 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', `
                 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade gesetzt: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

$installedVersion = $null
if ($Path) {
    $installedVersion = ($Path | Select-Object -First 1).DisplayVersion
    Write_LogEntry -Message "Gefundene installierte Version (Registry): $($installedVersion)" -Level "DEBUG"
}

$Install = $false
if ($installedVersion) {
    try {
        if ($installerPath) {
            $localVerForCompare = GetLocalVersionNumber -Path $installerPath
        } else {
            $localVerForCompare = $localVersion
        }
        if ($localVerForCompare -and ([version]$installedVersion -lt [version]$localVerForCompare)) {
            $Install = $true
        }
    } catch {
        if ($installedVersion -ne $localVersion) { $Install = $true }
    }
}

# Install if requested or required
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt; rufe Installationsskript mit Flag auf." -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\NotepadPlusPlusInstallation.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Installationsskript mit Flag aufgerufen." -Level "DEBUG"
} elseif ($Install) {
    Write_LogEntry -Message "Starte externes Installationsskript (Update) ohne Flag." -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\NotepadPlusPlusInstallation.ps1"
    Write_LogEntry -Message "Externes Installationsskript aufgerufen (Update)." -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===

