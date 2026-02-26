param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "PowerShell 7"
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

$localFileFilter = "PowerShell-*-win-x64.msi"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1"
$taskScript      = "$Serverip\Daten\Customize_Windows\Scripte\Aufgabenplannung_powershell_to_pwsh.ps1"

# ── Local version (from filename) ─────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'PowerShell-([\d.]+)-win-x64\.msi')
    if ($m.Success) { $localVersion = $m.Groups[1].Value }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
    $localVersion = "0.0.0"
}

# ── Online version via GitHub ──────────────────────────────────────────────────
$githubInfo = Get-GitHubLatestRelease `
    -Repo        "PowerShell/PowerShell" `
    -Token       $GitHubToken `
    -AssetFilter { param($a) $a.name -like "*-win-x64.msi" } `
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
    $needDownload = $false
    try {
        $needDownload = [version]$onlineVersion -gt [version]$localVersion
    } catch { $needDownload = $onlineVersion -ne $localVersion }

    if ($needDownload) {
        $destPath = Join-Path $InstallationFolder $assetName
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
}

Write-Host ""

# ── Re-evaluate local file ─────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = "0.0.0"
if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'PowerShell-([\d.]+)-win-x64\.msi')
    if ($m.Success) { $localVersion = $m.Groups[1].Value }
}

# ── Installed vs. local ────────────────────────────────────────────────────────
# PowerShell 7 registry DisplayVersion may have 4 parts — trim to 3
$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = $null
if ($installedInfo -and $installedInfo.VersionRaw) {
    $p = ($installedInfo.VersionRaw -split '\.') | Select-Object -First 3
    $installedVersion = $p -join '.'
}
$Install = $false

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

# ── Always update task scheduler entries powershell.exe → pwsh.exe ────────────
Write-Host ""
Write-Host "Ändere Powershell.exe zu Pwsh.exe in allen Tasks, wenn PS7 installiert ist." -ForegroundColor Magenta
Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $taskScript | Out-Null

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
