param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WireGuard"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$NetworkShareDaten  = $config.NetworkShareDaten

$InstallationFolder = "$InstallationFolder\WireGuard"
$localFileFilter    = "wireguard-amd64-*.msi"
$fileNameRegex      = 'wireguard-amd64-(\d+\.\d+\.\d+)\.msi'
$installScript      = "$NetworkShareDaten\Prog\InstallationScripts\Installation\WireGuardInstall.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $localVersion = Get-InstallerFileVersion -FilePath $localFile.FullName -FileNameRegex $fileNameRegex -Source FileName
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale WireGuard-Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version from GitHub tags ───────────────────────────────────────────
$onlineVersion = $null
try {
    $html          = Invoke-RestMethod -Uri 'https://github.com/WireGuard/wireguard-windows/tags' -UseBasicParsing -ErrorAction Stop
    $onlineVersion = [regex]::Matches($html, '/WireGuard/wireguard-windows/releases/tag/v(\d+\.\d+\.\d+)') |
        ForEach-Object { $_.Groups[1].Value } |
        Sort-Object { [version]$_ } -Descending |
        Select-Object -First 1
    Write-DeployLog -Message "Online-Version: $onlineVersion" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Online-Version: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $onlineVersion -and $localVersion) {
    $isNewer = $false
    try { $isNewer = [version]$onlineVersion -gt [version]$localVersion } catch { $isNewer = $onlineVersion -ne $localVersion }

    if ($isNewer) {
        $destPath    = Join-Path $InstallationFolder "wireguard-amd64-$onlineVersion.msi"
        $downloadUrl = "https://download.wireguard.com/windows-client/wireguard-amd64-$onlineVersion.msi"

        [void](Invoke-InstallerDownload `
            -Url                $downloadUrl `
            -OutFile            $destPath `
            -ConfirmDownload `
            -ReplaceOld `
            -RemoveFiles        @($localFile.FullName) `
            -KeepFiles          @($destPath) `
            -EmitHostStatus `
            -SuccessHostMessage "$ProgramName wurde aktualisiert.." `
            -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
            -Context            $ProgramName)

        # Re-read local file after potential update
        $localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
        $localVersion = if ($localFile) { Get-InstallerFileVersion -FilePath $localFile.FullName -FileNameRegex $fileNameRegex -Source FileName } else { $null }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFile) {
    Write-DeployLog -Message "Keine lokale Installationsdatei – Update-Vergleich übersprungen." -Level 'WARNING'
} elseif (-not $onlineVersion) {
    Write-DeployLog -Message "Online-Version konnte nicht ermittelt werden." -Level 'WARNING'
}

Write-Host ""

# ── Installed vs. local ───────────────────────────────────────────────────────
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

# ── Install if needed ─────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
