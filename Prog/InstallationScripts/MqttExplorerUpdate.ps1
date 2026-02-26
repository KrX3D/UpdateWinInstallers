param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "MQTT Explorer"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$GitHubToken        = $config.GitHubToken

$localFileFilter = "MQTT-Explorer*.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\MqttExplorerInstallation.ps1"
$includeBeta     = $true   # set $false to only track stable releases

# ── Helper: parse version string including optional beta suffix ───────────────
function Get-MqttVersionParts ([string]$Raw) {
    $clean = ($Raw -replace '^[vV]', '').Trim()
    # handle "-beta.6", "-beta6", etc.
    $beta = if ($clean -match '-beta\.?(\d+)') { [int]$Matches[1] } else { $null }
    $ver  = $clean -replace '-beta\.?\d+', '' -replace '-.*$', ''
    # Store cleaned display string (no leading 'v', proper beta format)
    $display = if ($beta) { "$ver-beta.$beta" } else { $ver }
    return @{ Version = $ver; Beta = $beta; Raw = $display }
}

function Compare-MqttVersions ([hashtable]$a, [hashtable]$b) {
    # Returns $true if $a > $b
    try {
        $vA = [version]$a.Version
        $vB = [version]$b.Version
        if ($vA -gt $vB) { return $true }
        if ($vA -lt $vB) { return $false }
        # same base version: final > any beta
        if (-not $a.Beta -and $b.Beta) { return $true }   # a is final, b is beta
        if ($a.Beta -and -not $b.Beta) { return $false }  # a is beta, b is final
        if ($a.Beta -and $b.Beta) { return [int]$a.Beta -gt [int]$b.Beta }
        return $false
    } catch { return $false }
}

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localParts   = @{ Version = '0.0.0'; Beta = $null; Raw = '0.0.0' }

if ($localFile) {
    $vi  = $localFile.VersionInfo
    $raw = if ($vi.ProductVersion) { $vi.ProductVersion } else { $vi.FileVersion }
    $localParts = Get-MqttVersionParts -Raw ($raw -replace '^[vV]', '').Trim()
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $($localParts.Raw)" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ─────────────────────────────────────────────────
$apiHeaders = @{
    'User-Agent' = 'InstallationScripts/1.0'
    'Accept'     = 'application/vnd.github.v3+json'
}
if ($GitHubToken) { $apiHeaders['Authorization'] = "token $GitHubToken" }

$chosenRelease = $null
$onlineParts   = $null

try {
    if ($includeBeta) {
        # Fetch all releases, find the one with the highest version
        $releases = Invoke-RestMethod -Uri 'https://api.github.com/repos/thomasnordquist/MQTT-Explorer/releases' -Headers $apiHeaders -ErrorAction Stop
        if ($releases -is [array] -and $releases.Count -gt 0) {
            $best = $null
            $bestParts = $null
            foreach ($rel in ($releases | Where-Object { -not $_.draft })) {
                $parts = Get-MqttVersionParts -Raw $rel.tag_name
                if (-not $best -or (Compare-MqttVersions -a $parts -b $bestParts)) {
                    $best = $rel
                    $bestParts = $parts
                }
            }
            $chosenRelease = $best
            $onlineParts   = $bestParts
        }
    } else {
        $release     = Invoke-RestMethod -Uri 'https://api.github.com/repos/thomasnordquist/MQTT-Explorer/releases/latest' -Headers $apiHeaders -ErrorAction Stop
        $chosenRelease = $release
        $onlineParts   = Get-MqttVersionParts -Raw $release.tag_name
    }
    if ($onlineParts) {
        Write-DeployLog -Message "Online-Version: $($onlineParts.Raw)" -Level 'INFO'
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der GitHub API: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $($localParts.Raw)"  -ForegroundColor Cyan
Write-Host "Online Version: $(if($onlineParts){ $onlineParts.Raw }else{'(unbekannt)'})" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($onlineParts -and $chosenRelease -and (Compare-MqttVersions -a $onlineParts -b $localParts)) {
    Write-DeployLog -Message "Neue Version verfügbar. Starte Download." -Level 'INFO'

    $asset = $chosenRelease.assets |
        Where-Object { $_.name -match '\.exe$' -and $_.name -notmatch '(?i)(portable|zip|tar|\.7z|\.gz)' -and $_.name -match '(?i)(setup|installer|mqtt)' } |
        Select-Object -First 1
    if (-not $asset) {
        $asset = $chosenRelease.assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1
    }

    if ($asset) {
        $fileName     = $asset.name
        $downloadPath = Join-Path $InstallationFolder $fileName
        $tempPath     = "$downloadPath.part"

        $ok = Invoke-DownloadFile -Url $asset.browser_download_url -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            try {
                Move-Item -Path $tempPath -Destination $downloadPath -Force -ErrorAction Stop
                if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $downloadPath) {
                    Remove-PathSafe -Path $localFile.FullName | Out-Null
                }
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert: $downloadPath" -Level 'SUCCESS'
                $localFile = Get-Item $downloadPath -ErrorAction SilentlyContinue
            } catch {
                Write-DeployLog -Message "Fehler beim Finalisieren: $($_.Exception.Message)" -Level 'ERROR'
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            }
        } else {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
        }
    } else {
        Write-DeployLog -Message "Kein passendes Asset gefunden." -Level 'WARNING'
    }
} elseif ($onlineParts) {
    Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile  = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localParts = @{ Version = '0.0.0'; Beta = $null; Raw = '0.0.0' }
if ($localFile) {
    $vi  = $localFile.VersionInfo
    $raw = if ($vi.ProductVersion) { $vi.ProductVersion } else { $vi.FileVersion }
    $localParts = Get-MqttVersionParts -Raw ($raw -replace '^[vV]', '').Trim()
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    $instParts = Get-MqttVersionParts -Raw $installedVersion

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $($instParts.Raw)"  -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $($localParts.Raw)" -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $($instParts.Raw) | Lokal: $($localParts.Raw)" -Level 'INFO'

    $Install = Compare-MqttVersions -a $localParts -b $instParts
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
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
