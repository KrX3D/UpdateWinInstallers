param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Microsoft Windows Desktop Runtime"
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

$InstallFolder      = "$InstallationFolder\ImageGlass"
$localFileFilter    = "windowsdesktop-runtime-*-win-x64.exe"
$targetMajorVersion = "8.0"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\WindowsDesktopRuntimesInstall.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallFolder -Filter $localFileFilter
$localVersion = "0.0.0"

if ($localFile) {
    $rawPV = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    if ($rawPV) { $localVersion = ($rawPV -split '\.' | Select-Object -First 3) -join '.' }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ──────────────────────────────────────────────────
$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'application/vnd.github.v3+json' }
if ($GitHubToken) { $headers['Authorization'] = "token $GitHubToken" }

$onlineVersion = $null
$downloadUrl   = $null

try {
    $releases = Invoke-RestMethod -Uri "https://api.github.com/repos/dotnet/windowsdesktop/releases" -Headers $headers -ErrorAction Stop
    $latest   = $releases |
        Where-Object { $_.tag_name -like "v$targetMajorVersion.*" -and -not $_.prerelease -and -not $_.draft } |
        Sort-Object { [version]($_.tag_name -replace '^v', '') } -Descending |
        Select-Object -First 1

    if ($latest) {
        $onlineVersion = $latest.tag_name -replace '^v', ''

        # Prefer a direct release asset
        $asset = $latest.assets | Where-Object { $_.name -match 'windowsdesktop.*runtime.*win.*x64.*\.exe$' } | Select-Object -First 1
        if ($asset) { $downloadUrl = $asset.browser_download_url }

        Write-DeployLog -Message "Online-Version: $onlineVersion | Asset: $($asset?.name)" -Level 'INFO'
    } else {
        Write-DeployLog -Message "Kein passender Release für v$targetMajorVersion gefunden." -Level 'WARNING'
    }
} catch {
    Write-DeployLog -Message "GitHub API Fehler: $_" -Level 'ERROR'
}

# Fallback: scrape the official dotnet download page for a direct link
if ($onlineVersion -and -not $downloadUrl) {
    try {
        $page = (Invoke-WebRequest `
            -Uri     "https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-$onlineVersion-windows-x64-installer" `
            -Headers @{ 'User-Agent' = 'InstallationScripts/1.0' } `
            -UseBasicParsing -ErrorAction Stop).Content
        $m = [regex]::Match($page, 'https?://[^"''<>\s]+windowsdesktop[^"''<>\s]*win[^"''<>\s]*x64[^"''<>\s]*\.exe', 'IgnoreCase')
        if ($m.Success) {
            $downloadUrl = $m.Value
            Write-DeployLog -Message "Fallback Download-URL: $downloadUrl" -Level 'DEBUG'
        } else {
            Write-DeployLog -Message "Kein direkter Download-Link auf der Dotnet-Seite gefunden." -Level 'WARNING'
        }
    } catch {
        Write-DeployLog -Message "Fehler beim Abrufen der Dotnet-Download-Seite: $_" -Level 'ERROR'
    }
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($onlineVersion -and $downloadUrl) {
    $isNewer = $false
    try { $isNewer = [version]$onlineVersion -gt [version]$localVersion } catch { $isNewer = $onlineVersion -ne $localVersion }

    if ($isNewer) {
        $destPath = Join-Path $InstallFolder "windowsdesktop-runtime-$onlineVersion-win-x64.exe"
        $tempPath = "$destPath.part"

        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            Move-Item -Path $tempPath -Destination $destPath -Force
            if ($localFile -and (Test-Path $localFile.FullName) -and $localFile.FullName -ne $destPath) {
                Remove-PathSafe -Path $localFile.FullName | Out-Null
            }
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
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
} else {
    Write-DeployLog -Message "Download-URL konnte nicht ermittelt werden." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallFolder -Filter $localFileFilter
$localVersion = "0.0.0"
if ($localFile) {
    $rawPV = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    if ($rawPV) { $localVersion = ($rawPV -split '\.' | Select-Object -First 3) -join '.' }
}

# ── Installed vs. local ───────────────────────────────────────────────────────
# .NET Desktop Runtime registers with the version in DisplayName (e.g. "... - 8.0.13 (x64)"),
# so we scan all matching entries and pick the highest version ourselves.
$regPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
)
$installedCandidates = foreach ($regPath in $regPaths) {
    if (Test-Path $regPath) {
        Get-ChildItem $regPath -ErrorAction SilentlyContinue |
            Get-ItemProperty -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "$ProgramName*" } |
            ForEach-Object {
                # Extract from DisplayName first (e.g. "...Runtime 8.0.13 (x64)");
                # DisplayVersion for .NET runtimes contains an internal build number, not semver.
                if ($_.DisplayName -match '([\d]+\.[\d]+(?:\.[\d]+)?)') {
                    $Matches[1]
                } elseif ($_.PSObject.Properties['DisplayVersion'] -and $_.DisplayVersion) {
                    $_.DisplayVersion
                }
            }
    }
}
$installedVersion = if ($installedCandidates) {
    ($installedCandidates | Sort-Object { [version]($_ -replace '[^\d\.]', '') } -Descending)[0]
} else { $null }

$Install = $false
if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    try { $Install = [version]($installedVersion -replace '[^\d\.]', '') -lt [version]($localVersion -replace '[^\d\.]', '') } catch { $Install = $false }
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
