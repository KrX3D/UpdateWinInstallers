param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Prismatik"
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

$localFileFilter = "Prismatik.unofficial.64bit.Setup*.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\PrismatikInstallation.ps1"

# ── Local version (from ProductVersion) ───────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = "0.0.0"

if ($localFile) {
    $rawPV = Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    if ($rawPV) { $localVersion = $rawPV }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ──────────────────────────────────────────────────
$githubInfo = Get-GitHubLatestRelease `
    -Repo        "psieg/Lightpack" `
    -Token       $GitHubToken `
    -AssetFilter { param($a) $a.name -match '\.exe$' -and ($a.name -match '64' -or $a.name -match 'unofficial') } `
    -Context     $ProgramName

$onlineVersion = if ($githubInfo) { $githubInfo.Version } else { $null }
$downloadUrl   = if ($githubInfo) { $githubInfo.DownloadUrl } else { $null }
$assetName     = if ($githubInfo) { $githubInfo.AssetName } else { $null }

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($onlineVersion -and $downloadUrl) {
    $isNewer = $false
    try { $isNewer = [version]$onlineVersion -gt [version]$localVersion } catch { $isNewer = $onlineVersion -ne $localVersion }

    if ($isNewer) {
        $destPath = Join-Path $InstallationFolder $assetName
        $tempPath = "$destPath.part"

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            Move-Item -Path $tempPath -Destination $destPath -Force
            if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $destPath) {
                Remove-PathSafe -Path $localFile.FullName | Out-Null
            }
            Write-Host "$ProgramName wurde aktualisiert: $assetName" -ForegroundColor Green
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
$localVersion = "0.0.0"
if ($localFile) {
    $rawPV = Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    if ($rawPV) { $localVersion = $rawPV }
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
