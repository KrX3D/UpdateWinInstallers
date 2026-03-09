param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "PrusaSlicer"
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

$InstallationFolder = Join-Path $InstallationFolder "3D"
$localFileFilter    = "prusaslicer_*.exe"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder" -Level 'INFO'

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
    -Repo         "prusa3d/PrusaSlicer" `
    -Token        $GitHubToken `
    -VersionRegex '(\d+\.\d+\.\d+)' `
    -Context      $ProgramName

# GitHub tag is like "version_2.9.1" — clean it
$onlineVersion = $null
if ($githubInfo -and $githubInfo.Tag) {
    $onlineVersion = ($githubInfo.Tag -replace '^version_', '') -replace '^[vV]', ''
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer — scrape actual installer URL from help.prusa3d.com ─────
if ($onlineVersion) {
    $isNewer = $false
    try { $isNewer = [version]$onlineVersion -gt [version]$localVersion } catch { $isNewer = $onlineVersion -ne $localVersion }

    if ($isNewer) {
        Write-DeployLog -Message "Neue Version: $onlineVersion. Suche Download-URL..." -Level 'INFO'

        $downloadUrl = $null
        try {
            $html = (Invoke-WebRequest -Uri 'https://help.prusa3d.com/downloads' -UseBasicParsing -ErrorAction Stop).Content

            # Normalize escaped JSON fragments and extract platform/file_url tuples.
            $normalized = $html -replace '\\/', '/' -replace '\\"', '"'
            $pattern = '"platform":"(win|standalone)".*?"file_url":"(https://[^"]+)"'
            $matches2 = [regex]::Matches($normalized, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

            $winVersions = @{}
            $standaloneVersions = @{}

            foreach ($match in $matches2) {
                $platform = $match.Groups[1].Value
                $link     = $match.Groups[2].Value -replace '(\.exe.*)$', '.exe'
                $lLower   = $link.ToLower()

                # Skip non-exe/archives/firmware and old snapshots.
                if ($lLower -notmatch '\.exe$' -or $lLower -match 'firmware|\.zip$|\.tar|\.gz$') { continue }
                if ($lLower -like '*old*') { continue }
                if ($lLower -notmatch 'prusaslicer_win') { continue }

                $vm = [regex]::Match($link, '(\d+\.\d+\.\d+|\d+_\d+_\d+)')
                if (-not $vm.Success) { continue }
                $ver = $vm.Groups[1].Value -replace '_', '.'

                if ($platform -eq 'win') {
                    $winVersions[$ver] = $link
                } elseif ($platform -eq 'standalone') {
                    $standaloneVersions[$ver] = $link
                }
            }

            $latestWinVersion = $winVersions.Keys | Sort-Object { [version]$_ } | Select-Object -Last 1
            $latestStandaloneVersion = $standaloneVersions.Keys | Sort-Object { [version]$_ } | Select-Object -Last 1

            if ($latestWinVersion -and $latestStandaloneVersion) {
                if ([version]$latestWinVersion -eq [version]$latestStandaloneVersion) {
                    # Prefer standalone when both are available with same version.
                    $bestVersion = $latestStandaloneVersion
                    $downloadUrl = $standaloneVersions[$latestStandaloneVersion]
                } elseif ([version]$latestWinVersion -gt [version]$latestStandaloneVersion) {
                    $bestVersion = $latestWinVersion
                    $downloadUrl = $winVersions[$latestWinVersion]
                } else {
                    $bestVersion = $latestStandaloneVersion
                    $downloadUrl = $standaloneVersions[$latestStandaloneVersion]
                }
            } elseif ($latestStandaloneVersion) {
                $bestVersion = $latestStandaloneVersion
                $downloadUrl = $standaloneVersions[$latestStandaloneVersion]
            } elseif ($latestWinVersion) {
                $bestVersion = $latestWinVersion
                $downloadUrl = $winVersions[$latestWinVersion]
            }

            # GitHub decides if update exists; URL must match that version.
            if ($downloadUrl -and $onlineVersion -and $bestVersion -ne $onlineVersion) {
                Write-DeployLog -Message "Versionen stimmen nicht überein. Gewünscht: $onlineVersion, URL: $bestVersion" -Level 'WARNING'
                $downloadUrl = $null
            }

            if ($downloadUrl) {
                Write-DeployLog -Message "Download-URL: $downloadUrl (Version: $bestVersion)" -Level 'INFO'
            } else {
                Write-DeployLog -Message "Keine passende Download-URL für Online-Version $onlineVersion gefunden." -Level 'WARNING'
            }
        } catch {
            Write-DeployLog -Message "Fehler beim Abrufen der Download-Seite: $_" -Level 'ERROR'
        }

        if ($downloadUrl) {
            $newInstaller = Join-Path $InstallationFolder "prusaslicer_$onlineVersion.exe"
            $tempPath     = "$newInstaller.part"

            $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
            if ($ok -and (Test-Path $tempPath)) {
                Move-Item -Path $tempPath -Destination $newInstaller -Force
                if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $newInstaller) {
                    Remove-PathSafe -Path $localFile.FullName | Out-Null
                }
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert: $newInstaller" -Level 'SUCCESS'
                $localFile = Get-Item $newInstaller -ErrorAction SilentlyContinue
            } else {
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
                Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
            }
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} else {
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