param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Hass.Agent"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config      = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$Serverip    = $config.Serverip
$PSHostPath  = $config.PSHostPath
$GitHubToken = $config.GitHubToken
# Dot-source config to expose $NetworkShareDaten and other extended variables
. (Get-SharedConfigPath -ScriptRoot $PSScriptRoot)

$InstallationFolder = "$NetworkShareDaten\Projekte\Smart_Home\HASS_Agent"
$localFileFilter    = "HASS.Agent.Installer.exe"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\HassAgentInstallation.ps1"
$includeBeta        = $true

Write-DeployLog -Message "InstallationFolder: $InstallationFolder | IncludeBeta: $includeBeta" -Level 'INFO'

# ── Helper: parse version + optional beta suffix ──────────────────────────────
function Get-VersionParts ([string]$Raw) {
    $base = ($Raw -split ' ')[0] -replace '^[vV]', ''
    # Handle both "-beta6" and "-beta.6" formats
    $beta = if ($base -match '-beta\.?(\d+)') { [int]$Matches[1] } else { $null }
    $ver  = $base -replace '-beta\.?\d+', ''
    return @{ Version = $ver; Beta = $beta }
}

# ── Local version ─────────────────────────────────────────────────────────────
$localFile  = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localParts = @{ Version = '0.0.0'; Beta = $null }

if ($localFile) {
    try {
        $raw        = ((Get-Item $localFile.FullName).VersionInfo.ProductVersion -split ' ')[0]
        $localParts = Get-VersionParts -Raw $raw
    } catch {
        Write-DeployLog -Message "Fehler beim Lesen der lokalen Version: $_" -Level 'WARNING'
    }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $($localParts.Version) Beta: $($localParts.Beta)" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ─────────────────────────────────────────────────
$apiHeaders = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GitHubToken) { $apiHeaders['Authorization'] = "token $GitHubToken" }

$apiUrl = if ($includeBeta) {
    'https://api.github.com/repos/hass-agent/HASS.Agent/releases'
} else {
    'https://api.github.com/repos/hass-agent/HASS.Agent/releases/latest'
}

$chosenRelease = $null
try {
    $releases = Invoke-RestMethod -Uri $apiUrl -Headers $apiHeaders -ErrorAction Stop
    if ($includeBeta -and $releases -is [array]) {
        $chosenRelease = $releases |
            Where-Object { -not $_.draft } |
            Sort-Object { if ($_.published_at) { [datetime]$_.published_at } else { [datetime]::MinValue } } -Descending |
            Select-Object -First 1
    } else {
        $chosenRelease = $releases
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der GitHub API: $_" -Level 'ERROR'
}

$onlineParts = $null
if ($chosenRelease) {
    $onlineParts = Get-VersionParts -Raw $chosenRelease.tag_name
    Write-DeployLog -Message "Online-Version: $($onlineParts.Version) Beta: $($onlineParts.Beta)" -Level 'INFO'
}

Write-Host ""
Write-Host "Lokale Version: $($localParts.Version)$(if($localParts.Beta){" Beta$($localParts.Beta)"})" -ForegroundColor Cyan
Write-Host "Online Version: $(if($onlineParts){"$($onlineParts.Version)$(if($onlineParts.Beta){" Beta$($onlineParts.Beta)"})"}else{"(unbekannt)"})" -ForegroundColor Cyan
Write-Host ""

# ── Determine if update needed ────────────────────────────────────────────────
$needUpdate = $false
if ($onlineParts) {
    try {
        $vO = [version]$onlineParts.Version
        $vL = [version]$localParts.Version
        if ($vO -gt $vL) { $needUpdate = $true }
        elseif ($vO -eq $vL) {
            # online is final release, local is a beta → online is newer
            if (-not $onlineParts.Beta -and $localParts.Beta) { $needUpdate = $true }
            # both beta but online has higher number
            elseif ($onlineParts.Beta -and $localParts.Beta -and [int]$onlineParts.Beta -gt [int]$localParts.Beta) { $needUpdate = $true }
        }
    } catch { if ($onlineParts.Version -ne $localParts.Version) { $needUpdate = $true } }
}

# ── Download if needed ────────────────────────────────────────────────────────
if ($needUpdate -and $chosenRelease) {
    Write-DeployLog -Message "Update erkannt: $($onlineParts.Version) > $($localParts.Version)" -Level 'INFO'

    $asset = $chosenRelease.assets |
        Where-Object { $_.name -match 'Installer' -and $_.name -match '\.exe$' } |
        Select-Object -First 1
    if (-not $asset) {
        $asset = $chosenRelease.assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1
    }

    if ($asset) {
        $newFilePath = Join-Path $InstallationFolder $asset.name
        $tempPath    = "$newFilePath.part"

        $ok = Invoke-DownloadFile -Url $asset.browser_download_url -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            Move-Item -Path $tempPath -Destination $newFilePath -Force
            if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $newFilePath) {
                Remove-PathSafe -Path $localFile.FullName | Out-Null
            }
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert: $newFilePath" -Level 'SUCCESS'
            $localFile = Get-Item $newFilePath -ErrorAction SilentlyContinue
        } else {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
        }
    } else {
        Write-DeployLog -Message "Kein passendes Asset gefunden." -Level 'WARNING'
    }
} elseif (-not $needUpdate) {
    Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
}

Write-Host ""

# ── Re-read local version ─────────────────────────────────────────────────────
$localFile  = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localParts = @{ Version = '0.0.0'; Beta = $null }
if ($localFile) {
    try {
        $raw        = ((Get-Item $localFile.FullName).VersionInfo.ProductVersion -split ' ')[0]
        $localParts = Get-VersionParts -Raw $raw
    } catch {}
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    $instParts = Get-VersionParts -Raw $installedVersion

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $($localParts.Version)$(if($localParts.Beta){" Beta$($localParts.Beta)"})" -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $($localParts.Version)" -Level 'INFO'

    try {
        $vI = [version]$instParts.Version
        $vL = [version]$localParts.Version
        if ($vI -lt $vL) { $Install = $true }
        elseif ($vI -eq $vL) {
            if ($instParts.Beta -and -not $localParts.Beta) { $Install = $true }
            elseif ($instParts.Beta -and $localParts.Beta -and [int]$instParts.Beta -lt [int]$localParts.Beta) { $Install = $true }
        }
    } catch { if ($instParts.Version -ne $localParts.Version) { $Install = $true } }

    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht in Registry gefunden." -Level 'INFO'
}

Write-Host ""

# ── Install if needed ─────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
