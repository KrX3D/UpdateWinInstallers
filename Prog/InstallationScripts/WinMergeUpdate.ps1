param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinMerge"
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

$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\WinMergeInstallation.ps1"
$localFileFilter = "WinMerge-*.exe"

# Local version (FileVersion from installer exe)
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null
if ($localFile) {
    $localVersion = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Kein lokaler WinMerge-Installer gefunden." -Level 'WARNING'
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
    return
}

# Online version via GitHub (list releases, filter to stable main releases only)
$githubInfo   = Get-GitHubLatestRelease -Repo "WinMerge/winmerge" -Token $GitHubToken `
    -AssetFilter { $_.name -match 'x64-Setup\.exe$' -and $_.name -notmatch 'ARM64|PerUser' } `
    -VersionRegex '(\d+\.\d+\.\d+(?:\.\d+)?)'
$onlineVersion = if ($githubInfo) { $githubInfo.Version } else { $null }
$downloadUrl   = if ($githubInfo) { $githubInfo.DownloadUrl } else { $null }

# Asset selection fallback (broader criteria if strict match missed)
if ($githubInfo -and -not $downloadUrl) {
    $downloadUrl = $githubInfo.DownloadUrl
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# Download if newer
if ($onlineVersion) {
    $needDownload = $false
    try {
        $lv = [version]($localVersion  -replace '[^\d\.]','')
        $ov = [version]($onlineVersion -replace '[^\d\.]','')
        $needDownload = $ov -gt $lv
    } catch { $needDownload = $onlineVersion -ne $localVersion }

    if ($needDownload) {
        # Fallback URL if asset not resolved
        if (-not $downloadUrl) {
            $downloadUrl = "https://github.com/WinMerge/winmerge/releases/download/v$onlineVersion/WinMerge-$onlineVersion-x64-Setup.exe"
        }
        $filename    = [System.IO.Path]::GetFileName(([uri]$downloadUrl).AbsolutePath)
        $destPath    = Join-Path $InstallationFolder $filename
        $tempPath    = "$destPath.part"

        Write-DeployLog -Message "Download: $downloadUrl -> $destPath" -Level 'INFO'
        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            Move-Item -Path $tempPath -Destination $destPath -Force
            Remove-Item -Path $localFile.FullName -Force -ErrorAction SilentlyContinue
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert: $destPath" -Level 'SUCCESS'
            $localFile = Get-Item $destPath -ErrorAction SilentlyContinue
        } else {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfuegbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} else {
    Write-Host "Kein Online Update verfuegbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Online-Version nicht ermittelt." -Level 'WARNING'
}

Write-Host ""

# Re-evaluate local file and check installed version
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null
if ($localFile) {
    try { $localVersion = (Get-Item -LiteralPath $localFile.FullName).VersionInfo.ProductVersion } catch {}
}

$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'
    try {
        $Install = [version]$installedVersion -lt [version]$localVersion
    } catch { $Install = $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "$ProgramName nicht in Registry." -Level 'INFO'
}

Write-Host ""

# Install if needed
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
