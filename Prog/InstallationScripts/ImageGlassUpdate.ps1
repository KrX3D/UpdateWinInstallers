param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "ImageGlass"
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

$localFileFilter = "ImageGlass*.msi"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\ImageGlassInstallation.ps1"

# ── Local version (4-part version embedded in filename) ───────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = "0.0.0.0"

if ($localFile) {
    $extracted = [regex]::Match($localFile.Name, '(?<=_)\d+\.\d+\.\d+\.\d+').Value
    $localVersion = if ($extracted) { $extracted } else {
        Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ─────────────────────────────────────────────────
$githubInfo = Get-GitHubLatestRelease `
    -Repo        "d2phap/ImageGlass" `
    -Token       $GitHubToken `
    -AssetFilter { param($a) $a.name -notlike '*delete*' -and $a.name -like '*x64*' -and $a.name -like '*.msi' } `
    -VersionRegex '(\d+\.\d+\.\d+\.\d+)' `
    -Context     $ProgramName

$onlineVersion = $null
$downloadUrl   = $null

if ($githubInfo) {
    $onlineVersion = $githubInfo.Version
    $downloadUrl   = $githubInfo.DownloadUrl
    Write-DeployLog -Message "Online-Version: $onlineVersion | Asset: $($githubInfo.AssetName)" -Level 'INFO'
} else {
    Write-DeployLog -Message "GitHub API lieferte keine Release-Informationen." -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $onlineVersion -and $downloadUrl) {
    $doDownload = try { [version]$onlineVersion -gt [version]$localVersion } catch { $onlineVersion -ne $localVersion }

    if ($doDownload) {
        Write-DeployLog -Message "Neue Version verfügbar. Starte Download: $downloadUrl" -Level 'INFO'

        $fileName     = Split-Path $downloadUrl -Leaf
        $downloadPath = Join-Path $InstallationFolder $fileName
        $tempPath     = "$downloadPath.part"

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            try {
                Move-Item -Path $tempPath -Destination $downloadPath -Force -ErrorAction Stop
                if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $downloadPath) {
                    Remove-PathSafe -Path $localFile.FullName | Out-Null
                }
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert: $downloadPath" -Level 'SUCCESS'
            } catch {
                Write-DeployLog -Message "Fehler beim Finalisieren: $($_.Exception.Message)" -Level 'ERROR'
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            }
        } else {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen. Temp-Datei fehlt: $tempPath" -Level 'ERROR'
        }
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
$localVersion = "0.0.0.0"
if ($localFile) {
    $extracted    = [regex]::Match($localFile.Name, '(?<=_)\d+\.\d+\.\d+\.\d+').Value
    $localVersion = if ($extracted) { $extracted } else {
        Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    }
}

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
