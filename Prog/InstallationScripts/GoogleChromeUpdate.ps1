param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Google Chrome"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "GoogleChromeStandaloneEnterprise64_*.msi"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1"

# ── Local version (from filename) ─────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $localVersion = ($localFile.Name -split '_')[1] -replace '\.msi$'
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via Google API ─────────────────────────────────────────────
$onlineVersion = $null
try {
    $apiResp = Invoke-RestMethod -Uri 'https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions' -ErrorAction Stop
    if ($apiResp.versions -is [array] -and $apiResp.versions.Count -gt 0) {
        $onlineVersion = $apiResp.versions[0].version
        Write-DeployLog -Message "Online-Version: $onlineVersion" -Level 'INFO'
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Online-Version: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $onlineVersion) {
    $doDownload = try { [version]$onlineVersion -gt [version]$localVersion } catch { $false }

    if ($doDownload) {
        Write-DeployLog -Message "Neuere Version verfügbar: $onlineVersion > $localVersion" -Level 'INFO'

        $downloadLink = 'https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi'
        $destPath     = Join-Path $InstallationFolder "GoogleChromeStandaloneEnterprise64_$onlineVersion.msi"

        [void](Invoke-InstallerDownload `
            -Url                $downloadLink `
            -OutFile            $destPath `
            -ConfirmDownload `
            -ReplaceOld `
            -RemoveFiles        @($localFile.FullName) `
            -KeepFiles          @($destPath) `
            -EmitHostStatus `
            -SuccessHostMessage "$ProgramName wurde aktualisiert.." `
            -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
            -SuccessLogMessage  "$ProgramName erfolgreich aktualisiert: $destPath" `
            -FailureLogMessage  "Download fehlgeschlagen: $destPath" `
            -Context            $ProgramName)
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFile) {
    Write-DeployLog -Message "Keine lokale Installationsdatei – Update-Vergleich übersprungen." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = if ($localFile) { ($localFile.Name -split '_')[1] -replace '\.msi$' } else { $null }

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
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
