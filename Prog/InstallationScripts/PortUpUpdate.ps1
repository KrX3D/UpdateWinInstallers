param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "PortUp"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
# Dot-source config to get $NetworkShareDaten
. (Get-SharedConfigPath -ScriptRoot $PSScriptRoot)

$InstallationFolder = "$NetworkShareDaten\Customize_Windows"
$localFilePath      = "$InstallationFolder\Tools\PortableUpdate\PortUp.exe"
$destinationPath    = Join-Path $env:USERPROFILE "Desktop\Reducer\PortUp.exe"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder" -Level 'INFO'

# ── Local version (from FileVersion, 3 significant parts) ─────────────────────
$localVersion = $null
if (Test-Path $localFilePath) {
    $rawFV = Get-InstallerFileVersion -FilePath $localFilePath -Source FileVersion
    if ($rawFV) {
        $parts   = $rawFV -split '\.'
        $trimmed = ($parts[0], $parts[1], ($parts[2] -replace '^0+(\d)', '$1')) -join '.'
        $localVersion = $trimmed
    }
    Write-DeployLog -Message "Lokale Datei: PortUp.exe | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale PortUp.exe gefunden: $localFilePath" -Level 'WARNING'
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
    return
}

# ── Online version ─────────────────────────────────────────────────────────────
$onlineInfo = Get-OnlineVersionInfo `
    -Url     'https://www.portableupdate.com/download' `
    -Regex   @('Portable Update (\d+\.\d+\.\d+)') `
    -Context $ProgramName

$onlineVersion = $onlineInfo.Version
$downloadUrl   = if ($onlineVersion) { "https://file.portableupdate.com/downloads/PortUp_$onlineVersion.zip" } else { $null }

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download and extract if newer ─────────────────────────────────────────────
if ($onlineVersion -and $downloadUrl) {
    $needDownload = $false
    try {
        $needDownload = [version]$onlineVersion -gt [version]$localVersion
    } catch { $needDownload = $onlineVersion -ne $localVersion }

    if ($needDownload) {
        $zipFileName = "PortUp_$onlineVersion.zip"
        $tempZip     = Join-Path $env:TEMP $zipFileName

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempZip
        if ($ok -and (Test-Path $tempZip)) {
            try {
                Expand-Archive -Path $tempZip -DestinationPath $InstallationFolder -Force
                Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert auf Version $onlineVersion." -Level 'SUCCESS'

                # Re-read local version after update
                if (Test-Path $localFilePath) {
                    $rawFV2 = Get-InstallerFileVersion -FilePath $localFilePath -Source FileVersion
                    if ($rawFV2) {
                        $parts2 = $rawFV2 -split '\.'
                        $localVersion = ($parts2[0], $parts2[1], ($parts2[2] -replace '^0+(\d)', '$1')) -join '.'
                    }
                }
            } catch {
                Write-Host "Extraktion fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
                Write-DeployLog -Message "Extraktion fehlgeschlagen: $_" -Level 'ERROR'
                Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue
            }
        } else {
            Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue
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

# ── Installed (desktop) vs. source ────────────────────────────────────────────
# PortUp has no registry entry — compare source file vs desktop copy
$installedVersion = $null
if (Test-Path $destinationPath) {
    $rawDest = Get-InstallerFileVersion -FilePath $destinationPath -Source FileVersion
    if ($rawDest) {
        $dparts = $rawDest -split '\.'
        $installedVersion = ($dparts[0], $dparts[1], ($dparts[2] -replace '^0+(\d)', '$1')) -join '.'
    }
}

$Install = $false
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
    Write-DeployLog -Message "$ProgramName nicht am Zielpfad gefunden." -Level 'INFO'
}

Write-Host ""

# ── Copy if needed ─────────────────────────────────────────────────────────────
if ($Install -or $InstallationFlag) {
    Write-Host "PortUp wird kopiert" -ForegroundColor Cyan
    Write-DeployLog -Message "Kopiere: $localFilePath -> $destinationPath" -Level 'INFO'

    $destDir = Split-Path $destinationPath -Parent
    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }

    if (Test-Path $localFilePath) {
        Copy-Item -Path $localFilePath -Destination $destinationPath -Force
        Write-DeployLog -Message "Kopiert: $destinationPath" -Level 'SUCCESS'
    } else {
        Write-DeployLog -Message "Quelldatei nicht gefunden: $localFilePath" -Level 'ERROR'
    }
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
