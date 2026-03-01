param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Autodesk Fusion"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = Join-Path $config.InstallationFolder "3D"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localInstaller   = Join-Path $InstallationFolder 'Fusion_Client_Downloader.exe'
$downloadUrl      = 'https://dl.appstreaming.autodesk.com/production/installers/Fusion%20Client%20Downloader.exe'
$installScript    = "$Serverip\Daten\Prog\InstallationScripts\Installation\Fusion360Installation.ps1"

Write-DeployLog -Message "LocalInstaller: $localInstaller" -Level 'DEBUG'

# ── Local version ─────────────────────────────────────────────────────────────
# Fusion stores version as parts[1].parts[2].parts[0] in FileVersion
function Get-FusionVersion ([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    $raw = (Get-Item $Path).VersionInfo.FileVersion
    if (-not $raw) { return $null }
    $p = $raw.Split('.')
    if ($p.Length -lt 3) { return $null }
    return "$($p[1]).$($p[2]).$($p[0])"
}

$localVersion = Get-FusionVersion -Path $localInstaller

# ── Online version ────────────────────────────────────────────────────────────
$onlineVersion = $null
try {
    $resp = Invoke-WebRequest -Uri 'http://autode.sk/whatsnew' -UseBasicParsing -ErrorAction Stop
    $hits = [regex]::Matches($resp.Content, 'v\.(\d+\.\d+\.\d+)')
    if ($hits.Count -gt 0) {
        $onlineVersion = $hits | ForEach-Object { $_.Groups[1].Value } |
            Sort-Object { [version]$_ } -Descending | Select-Object -First 1
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Online-Version: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download and replace if newer ─────────────────────────────────────────────
$tempPath   = Join-Path $env:TEMP 'Fusion_Client_Downloader.exe'
$needUpdate = $false

if (-not $localVersion) {
    Write-DeployLog -Message "Lokale Version unbekannt – Download wird erzwungen." -Level 'WARNING'
    $needUpdate = $true
} elseif ($onlineVersion) {
    $needUpdate = try { [version]$onlineVersion -gt [version]$localVersion } catch { $true }
}

if ($needUpdate) {
    if ($onlineVersion) {
        Write-DeployLog -Message "Update: $onlineVersion > $localVersion" -Level 'INFO'
    }
    $ok = $false
    try {
        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
    } catch {
        Write-DeployLog -Message "Download-Fehler: $_" -Level 'ERROR'
    }

    if ($ok -and (Test-Path $tempPath)) {
        if (Test-Path $localInstaller) {
            try {
                Remove-Item $localInstaller -Force
                Move-Item -Path $tempPath -Destination $InstallationFolder -Force
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "Neue Datei verschoben nach $InstallationFolder" -Level 'SUCCESS'
            } catch {
                Write-DeployLog -Message "Fehler beim Ersetzen: $_" -Level 'ERROR'
                Write-Host "Fehler beim Ersetzen der Datei." -ForegroundColor Red
            }
        } else {
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Lokale Datei zum Ersetzen nicht vorhanden." -Level 'ERROR'
        }
    } else {
        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
        Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
    }
} else {
    Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
}

Write-Host ""

# ── Re-evaluate local version ─────────────────────────────────────────────────
$localVersion = Get-FusionVersion -Path $localInstaller

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    $Install = try { [version]$installedVersion -lt [version]$localVersion } catch { $false }
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
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -InstallationFlag
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
