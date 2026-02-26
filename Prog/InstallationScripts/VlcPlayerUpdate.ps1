param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "VLC"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "vlc-*-win64.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\VlcPlayerInstall.ps1"

# ── Local version (from filename) ─────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'vlc-([\d.]+)-win64\.exe')
    if ($m.Success) { $localVersion = $m.Groups[1].Value }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale VLC-Installationsdatei gefunden." -Level 'INFO'
}

# ── Online version ─────────────────────────────────────────────────────────────
$onlineVersion = $null
$downloadUrl   = $null

try {
    $html = (Invoke-WebRequest -Uri 'https://www.videolan.org/vlc/index.html' -UseBasicParsing -ErrorAction Stop).Content
    $m2   = [regex]::Match($html, 'vlc-([\d.]+)-win64\.exe')
    if ($m2.Success) {
        $onlineVersion = $m2.Groups[1].Value
        $downloadUrl   = "https://vlc.pixelx.de/vlc/$onlineVersion/win64/vlc-$onlineVersion-win64.exe"
    }
    Write-DeployLog -Message "Online-Version: $onlineVersion" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der VLC-Seite: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($onlineVersion -and $downloadUrl) {
    $needDownload = $false
    try {
        $needDownload = (-not $localVersion) -or ([version]$onlineVersion -gt [version]$localVersion)
    } catch { $needDownload = $onlineVersion -ne $localVersion }

    if ($needDownload) {
        $destPath = Join-Path $InstallationFolder "vlc-$onlineVersion-win64.exe"
        $tempPath = "$destPath.part"

        # Enable TLS 1.2
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if (-not $ok -or -not (Test-Path $tempPath)) {
            # Retry via http mirror
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            $httpUrl = "http://download.videolan.org/pub/videolan/vlc/$onlineVersion/win64/vlc-$onlineVersion-win64.exe"
            Write-DeployLog -Message "Retry via HTTP: $httpUrl" -Level 'WARNING'
            $ok = Invoke-DownloadFile -Url $httpUrl -OutFile $tempPath
        }

        if ($ok -and (Test-Path $tempPath)) {
            Move-Item -Path $tempPath -Destination $destPath -Force
            if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $destPath) {
                Remove-PathSafe -Path $localFile.FullName | Out-Null
            }
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert: $destPath" -Level 'SUCCESS'
            $localFile = Get-Item $destPath -ErrorAction SilentlyContinue
        } else {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $onlineVersion) {
    Write-DeployLog -Message "Online-Version konnte nicht ermittelt werden." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ─────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null
if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'vlc-([\d.]+)-win64\.exe')
    if ($m.Success) { $localVersion = $m.Groups[1].Value }
}

# ── Installed vs. local ────────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    try { $Install = [version]$installedVersion -lt [version]$localVersion } catch { $Install = $false }
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

# ── Install if needed ──────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
