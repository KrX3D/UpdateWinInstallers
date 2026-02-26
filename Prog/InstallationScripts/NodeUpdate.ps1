param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Node.js"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "node-v*-x64.msi"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1"

# ── Local version (from filename) ─────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'node-v([\d.]+)-x64\.msi')
    if ($m.Success) { $localVersion = $m.Groups[1].Value }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version ─────────────────────────────────────────────────────────────
$onlineVersion = $null
$downloadUrl   = $null

$onlineInfo = Get-OnlineVersionInfo `
    -Url     'https://nodejs.org/download/release/latest/' `
    -Regex   @('href="[^"]*?(node-v([\d.]+)-x64\.msi)"') `
    -RegexGroup 2 `
    -Context $ProgramName

if ($onlineInfo.Content) {
    $m2 = [regex]::Match($onlineInfo.Content, 'href="[^"]*?(node-v([\d.]+)-x64\.msi)"')
    if ($m2.Success) {
        $onlineVersion = $m2.Groups[2].Value
        $downloadUrl   = "https://nodejs.org/download/release/latest/$($m2.Groups[1].Value)"
    }
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($onlineVersion -and $downloadUrl) {
    $needDownload = $false
    try {
        $needDownload = (-not $localVersion) -or ([version]$onlineVersion -gt [version]$localVersion)
    } catch { $needDownload = $onlineVersion -ne $localVersion }

    if ($needDownload) {
        $fileName  = "node-v$onlineVersion-x64.msi"
        $destPath  = Join-Path $InstallationFolder $fileName

        $ok = Invoke-InstallerDownload `
            -Url                $downloadUrl `
            -OutFile            $destPath `
            -ConfirmDownload `
            -ReplaceOld `
            -RemoveFiles        @($(if ($localFile) { $localFile.FullName } else { $null })) `
            -KeepFiles          @($destPath) `
            -EmitHostStatus `
            -SuccessHostMessage "$ProgramName wurde aktualisiert.." `
            -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
            -Context            $ProgramName
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
$localVersion = $null
if ($localFile) {
    $m = [regex]::Match($localFile.Name, 'node-v([\d.]+)-x64\.msi')
    if ($m.Success) { $localVersion = $m.Groups[1].Value }
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
