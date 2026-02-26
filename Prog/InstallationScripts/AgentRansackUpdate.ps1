param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Agent Ransack"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

# Wildcard matches both old .msi files and new .exe files
$localFileWildcard = "agentransack*"
$onlineVersionUrl  = "https://www.mythicsoft.com/agentransack/"
$installScript     = "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1"

# ── Local version (build number from filename) ────────────────────────────────
$localFileObj = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard |
                Where-Object { $_.Extension -match '\.(exe|msi)$' } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

$localFilePath = if ($localFileObj) { $localFileObj.FullName } else { $null }

$localVersion = if ($localFilePath) {
    [System.IO.Path]::GetFileNameWithoutExtension($localFilePath) `
        -replace 'agentransack_', '' -replace 'x64_', '' -replace 'msi_', ''
} else { $null }

Write-DeployLog -Message "Lokale Datei: $localFilePath | Version: $localVersion" -Level 'DEBUG'

# ── Online version + tokenized download link ──────────────────────────────────
# The download URL now contains a one-time/rotating token:
#   https://download.mythicsoft.com/flp/{build}/{token}/agentransack_{build}.exe
# So we scrape the actual href from the download button on the page.

$onlineVersion = $null
$downloadUrl   = $null

try {
    $html = (Invoke-WebRequest -Uri $onlineVersionUrl -UseBasicParsing -ErrorAction Stop).Content

    # Extract the tokenized exe download link
    $linkMatch = [regex]::Match($html,
        'href="(https://download\.mythicsoft\.com/flp/(\d+)/[^"]+\.exe)"')
    if ($linkMatch.Success) {
        $downloadUrl   = $linkMatch.Groups[1].Value
        $onlineVersion = $linkMatch.Groups[2].Value
        Write-DeployLog -Message "Online-Version: $onlineVersion | URL: $downloadUrl" -Level 'INFO'
    } else {
        # Fallback: extract build number from any agentransack reference
        $verMatch = [regex]::Match($html, 'agentransack[_-](\d{4,})')
        if ($verMatch.Success) {
            $onlineVersion = $verMatch.Groups[1].Value
            $downloadUrl   = "https://download.mythicsoft.com/flp/$onlineVersion/agentransack_$onlineVersion.exe"
            Write-DeployLog -Message "Fallback-URL konstruiert: $downloadUrl" -Level 'WARNING'
        }
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Webseite: $_" -Level 'ERROR'
}

# ── Compare and download ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

if ($localFilePath -and $onlineVersion -and $downloadUrl) {
    if ([int]$onlineVersion -gt [int]$localVersion) {
        $fileName  = Split-Path $downloadUrl -Leaf
        $destPath  = Join-Path $InstallationFolder $fileName

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $destPath
        if ($ok -and (Test-Path $destPath)) {
            if ($localFilePath -and (Test-Path $localFilePath) -and $localFilePath -ne $destPath) {
                Remove-PathSafe -Path $localFilePath | Out-Null
            }
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert auf Version $onlineVersion : $destPath" -Level 'SUCCESS'
            $localFilePath = $destPath
        } else {
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen: $destPath" -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFilePath) {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
} elseif (-not $onlineVersion) {
    Write-Host "Online-Version konnte nicht ermittelt werden." -ForegroundColor Red
    Write-DeployLog -Message "Online-Version nicht abrufbar." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFileObj  = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard |
                 Where-Object { $_.Extension -match '\.(exe|msi)$' } |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 1
$localFilePath = if ($localFileObj) { $localFileObj.FullName } else { $null }
$localVersion  = if ($localFilePath) {
    [System.IO.Path]::GetFileNameWithoutExtension($localFilePath) `
        -replace 'agentransack_', '' -replace 'x64_', '' -replace 'msi_', ''
} else { $null }

# ── Installed version (build number = 3rd segment of DisplayVersion "2.1.3555") ─
$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$Install       = $false

if ($installedInfo) {
    $installedVersion = ($installedInfo.VersionRaw -split '\.')[2]

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    if ([int]$installedVersion -lt [int]$localVersion) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
        $Install = $true
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
