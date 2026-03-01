param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Arduino"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config            = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = "$($config.InstallationFolder)\Arduino"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$GitHubToken        = $config.GitHubToken

$InstallationFilePattern = "arduino*.exe"
$installScript           = "$Serverip\Daten\Prog\InstallationScripts\Installation\ArduinoInstall.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
# ProductVersion (first 3 components), falling back to filename regex
$localInstaller = Get-InstallerFilePath -Directory $InstallationFolder -Filter $InstallationFilePattern -ExcludeNameLike "*_old*"
$localFileVersion = $null
if ($localInstaller) {
    $pv = Get-InstallerFileVersion -FilePath $localInstaller.FullName -Source ProductVersion
    $localFileVersion = if ($pv) {
        (($pv -split '\.')[0..2] -join '.')
    } else {
        Get-InstallerFileVersion -FilePath $localInstaller.FullName -FileNameRegex '(\d+\.\d+\.\d+)' -Source FileName
    }
}
if (-not $localFileVersion) { $localFileVersion = "0.0.0" }

Write-DeployLog -Message "Lokaler Installer: $($localInstaller.Name) | Version: $localFileVersion" -Level 'DEBUG'

# ── Online version via GitHub API ─────────────────────────────────────────────
$githubInfo = Get-GitHubLatestRelease `
    -Repo    "arduino/arduino-ide" `
    -Token   $GitHubToken `
    -AssetFilter {
        param($a)
        (($a.name -match '(?i)win|windows|x64|64bit|64-bit') -or
         ($a.browser_download_url -match '(?i)win|windows|x64|64bit|64-bit')) -and
        ($a.name -match '\.exe$')
    } `
    -VersionRegex '(\d+\.\d+\.\d+)' `
    -Context      $ProgramName

if ($githubInfo) {
    $onlineVersion    = $githubInfo.Version
    $downloadURL      = $githubInfo.DownloadUrl
    $downloadFileName = $githubInfo.AssetName

    Write-Host ""
    Write-Host "Lokale Version: $localFileVersion" -ForegroundColor Cyan
    Write-Host "Online Version: $onlineVersion"    -ForegroundColor Cyan
    Write-Host ""

    $doDownload = try { [version]$onlineVersion -gt [version]$localFileVersion } catch { $onlineVersion -ne $localFileVersion }

    if ($doDownload) {
        Write-DeployLog -Message "Neue Version verfügbar. Starte Download: $downloadURL" -Level 'INFO'

        # Download to a .part temp file, then move atomically (preserves original pattern)
        $downloadPath = Join-Path $InstallationFolder $downloadFileName
        $tempPath     = "$downloadPath.part"

        $ok = Invoke-DownloadFile -Url $downloadURL -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            try {
                Move-Item -Path $tempPath -Destination $downloadPath -Force -ErrorAction Stop
                Write-DeployLog -Message "Datei verschoben: $downloadPath" -Level 'DEBUG'

                if ($localInstaller -and (Test-Path -LiteralPath $localInstaller.FullName) -and
                    ($localInstaller.FullName -ne $downloadPath)) {
                    Remove-PathSafe -Path $localInstaller.FullName | Out-Null
                    Write-DeployLog -Message "Alte Installationsdatei entfernt: $($localInstaller.FullName)" -Level 'INFO'
                }

                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert: $downloadPath" -Level 'SUCCESS'
            } catch {
                Write-DeployLog -Message "Fehler beim Finalisieren: $($_.Exception.Message)" -Level 'ERROR'
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            }
        } else {
            Write-DeployLog -Message "Download fehlgeschlagen. Temp-Datei fehlt: $tempPath" -Level 'ERROR'
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "GitHub API lieferte keine Release-Informationen." -Level 'ERROR'
}

Write-Host ""

# ── Re-evaluate local file after potential download ───────────────────────────
$localInstaller = Get-InstallerFilePath -Directory $InstallationFolder -Filter $InstallationFilePattern -ExcludeNameLike "*_old*"
$localVersion   = $null
if ($localInstaller) {
    $pv = Get-InstallerFileVersion -FilePath $localInstaller.FullName -Source ProductVersion
    $localVersion = if ($pv) { (($pv -split '\.')[0..2] -join '.') } else {
        Get-InstallerFileVersion -FilePath $localInstaller.FullName -FileNameRegex '(\d+\.\d+\.\d+)' -Source FileName
    }
}

# ── Installed version vs. local installer ─────────────────────────────────────
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
