param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "NTLite"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$InstallationFolder = Join-Path $InstallationFolder "NTLite"
$localFilePath      = Join-Path $InstallationFolder "NTLite_setup_x64.exe"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder" -Level 'INFO'

# ── Local version (from FileVersion, 3 parts) ─────────────────────────────────
$localVersion = $null
if (Test-Path $localFilePath) {
    $rawFV = Get-InstallerFileVersion -FilePath $localFilePath -Source FileVersion
    if ($rawFV) {
        $parts = ($rawFV -split '\.') | Select-Object -First 3
        $localVersion = $parts -join '.'
    }
    Write-DeployLog -Message "Lokale Datei: NTLite_setup_x64.exe | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden: $localFilePath" -Level 'WARNING'
}

# ── Online version ─────────────────────────────────────────────────────────────
$onlineInfo = Get-OnlineVersionInfo `
    -Url     'https://www.ntlite.com/download/' `
    -Regex   @('v(\d+\.\d+\.\d+)') `
    -Context $ProgramName

$onlineVersion = $onlineInfo.Version
$downloadUrl   = $null

if ($onlineInfo.Content) {
    $lm = [regex]::Match($onlineInfo.Content, 'href="(https://downloads\.ntlite\.com/files/NTLite_setup_x64\.exe)"')
    if ($lm.Success) { $downloadUrl = $lm.Groups[1].Value }
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
        $tempPath = Join-Path $env:TEMP "NTLite_setup_x64.exe"

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            try {
                if (Test-Path $localFilePath) { Remove-Item -Path $localFilePath -Force }
                Move-Item -Path $tempPath -Destination $localFilePath -Force
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert: $localFilePath" -Level 'SUCCESS'
            } catch {
                Write-DeployLog -Message "Fehler beim Ersetzen der Datei: $_" -Level 'ERROR'
                Write-Host "Fehler beim Aktualisieren. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            }
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

# ── Re-evaluate local version ──────────────────────────────────────────────────
$localVersion = $null
if (Test-Path $localFilePath) {
    $rawFV = Get-InstallerFileVersion -FilePath $localFilePath -Source FileVersion
    if ($rawFV) {
        $parts = ($rawFV -split '\.') | Select-Object -First 3
        $localVersion = $parts -join '.'
    }
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
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -InstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
