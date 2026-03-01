param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Git"
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

$localFileFilter = "Git-*-64-bit.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\GitInstallation.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = "0.0.0"

if ($localFile) {
    $pv = Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    $localVersion = if ($pv) {
        (($pv -split '\.')[0..2] -join '.')
    } else {
        Get-InstallerFileVersion -FilePath $localFile.FullName -FileNameRegex 'Git-(\d+\.\d+\.\d+)' -Source FileName
    }
    if (-not $localVersion) { $localVersion = "0.0.0" }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ─────────────────────────────────────────────────
$githubInfo = Get-GitHubLatestRelease `
    -Repo        "git-for-windows/git" `
    -Token       $GitHubToken `
    -AssetFilter { param($a) $a.name -match 'Git-[\d.]+-64-bit\.exe$' } `
    -VersionRegex '(\d+\.\d+\.\d+)' `
    -Context     $ProgramName

$onlineVersion = $null
$downloadUrl   = $null

if ($githubInfo) {
    $onlineVersion = (($githubInfo.Version -split '\.')[0..2] -join '.')
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
$localVersion = "0.0.0"
if ($localFile) {
    $pv = Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    $localVersion = if ($pv) { (($pv -split '\.')[0..2] -join '.') } else {
        Get-InstallerFileVersion -FilePath $localFile.FullName -FileNameRegex 'Git-(\d+\.\d+\.\d+)' -Source FileName
    }
    if (-not $localVersion) { $localVersion = "0.0.0" }
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"    -ForegroundColor Cyan
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
