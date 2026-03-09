param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "EstlCam"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = Join-Path $config.InstallationFolder "CNC"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "Estlcam_64_*.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\EstlCamInstallation.ps1"

Write-DeployLog -Message "InstallationFlag: $InstallationFlag | InstallationFolder: $InstallationFolder" -Level 'INFO'

# ── Local version (integer build number from filename) ────────────────────────
$localFile  = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localBuild = $null

if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'Estlcam_64_(\d+)\.exe', 'IgnoreCase')
    if ($m.Success) { $localBuild = [int]$m.Groups[1].Value }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Build: $localBuild" -Level 'INFO'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version ────────────────────────────────────────────────────────────
$onlineInfo  = Get-OnlineVersionInfo -Url 'http://www.estlcam.de/download.htm' -Regex @('Estlcam_64_(\d+)\.exe') -Context $ProgramName
$onlineBuild = if ($onlineInfo.Version) { [int]$onlineInfo.Version } else { $null }

Write-Host ""
Write-Host "Lokale Version: $localBuild"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineBuild" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $null -ne $onlineBuild) {
    if ($onlineBuild -gt $localBuild) {
        Write-DeployLog -Message "Neuere Version verfügbar: $onlineBuild > $localBuild" -Level 'INFO'

        # Try to find a direct link on the page, else construct URL
        $pageInfo    = Get-OnlineInstallerLink `
            -Url         'http://www.estlcam.de/download.htm' `
            -LinkRegex   "href=""([^""]*Estlcam_64_$onlineBuild\.exe)""" `
            -Context     $ProgramName

        $downloadUrl = if ($pageInfo.DownloadUrl) {
            if ($pageInfo.DownloadUrl -notmatch '^https?://') { "http://www.estlcam.de/$($pageInfo.DownloadUrl.TrimStart('/'))" }
            else { $pageInfo.DownloadUrl }
        } else {
            "http://www.estlcam.de/Estlcam_64_$onlineBuild.exe"
        }

        $destPath = Join-Path $InstallationFolder "Estlcam_64_$onlineBuild.exe"

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $destPath
        if ($ok -and (Test-Path $destPath)) {
            if ($localFile -and (Test-Path $localFile.FullName)) {
                Remove-PathSafe -Path $localFile.FullName | Out-Null
            }
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert auf Build $onlineBuild." -Level 'SUCCESS'
        } else {
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen: $destPath" -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFile) {
    Write-DeployLog -Message "Keine lokale Installationsdatei – Update-Vergleich übersprungen." -Level 'WARNING'
} elseif ($null -eq $onlineBuild) {
    Write-Host "Online-Version konnte nicht ermittelt werden." -ForegroundColor Red
    Write-DeployLog -Message "Online-Version nicht verfügbar." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile  = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localBuild = $null
if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'Estlcam_64_(\d+)\.exe', 'IgnoreCase')
    if ($m.Success) { $localBuild = [int]$m.Groups[1].Value }
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    $installedBuild = try { [int]$installedVersion } catch { $null }

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedBuild" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localBuild"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedBuild | Lokal: $localBuild" -Level 'INFO'

    if ($null -ne $installedBuild -and $installedBuild -lt $localBuild) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
        $Install = $true
    } elseif ($installedBuild -eq $localBuild) {
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
