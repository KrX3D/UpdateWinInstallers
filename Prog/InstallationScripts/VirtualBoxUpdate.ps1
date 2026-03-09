param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Oracle VirtualBox"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$installerFilter  = "VirtualBox-*.exe"
$extPackFilter    = "Oracle_VirtualBox_Extension_Pack*.vbox*extpack"
$installScript    = "$Serverip\Daten\Prog\InstallationScripts\Installation\VirtualBoxInstall.ps1"

# ── Helper: component-by-component version comparison ─────────────────────────
function Compare-VersionComponents {
    param([string]$Local, [string]$Remote)
    $lp = $Local  -split '\.'
    $rp = $Remote -split '\.'
    $n  = [math]::Max($lp.Count, $rp.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $ln = 0
        $rn = 0
        if ($i -lt $lp.Count) { [void][int]::TryParse($lp[$i], [ref]$ln) }
        if ($i -lt $rp.Count) { [void][int]::TryParse($rp[$i], [ref]$rn) }
        $l = $ln
        $r = $rn
        if ($l -lt $r) { return $true }
        if ($l -gt $r) { return $false }
    }
    return $false
}

# ── Online version ─────────────────────────────────────────────────────────────
$onlineVersion     = $null
$installerFilename = $null
$extPackFilename   = $null

try {
    $onlineVersion = (Invoke-WebRequest -Uri 'https://download.virtualbox.org/virtualbox/LATEST.TXT' -UseBasicParsing -ErrorAction Stop).Content.Trim()
    Write-DeployLog -Message "Online-Version: $onlineVersion" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen von LATEST.TXT: $_" -Level 'ERROR'
}

if ($onlineVersion) {
    try {
        $md5 = (Invoke-WebRequest -Uri "https://download.virtualbox.org/virtualbox/$onlineVersion/MD5SUMS" -UseBasicParsing -ErrorAction Stop).Content
        $installerFilename = [regex]::Match($md5, 'VirtualBox-.*-Win\.exe').Value
        $extPackFilename   = [regex]::Match($md5, 'Oracle_VirtualBox_Extension_Pack-\d+\.\d+\.\d+-\d+\.vbox-extpack').Value
        Write-DeployLog -Message "Installer: $installerFilename | ExtPack: $extPackFilename" -Level 'DEBUG'
    } catch {
        Write-DeployLog -Message "Fehler beim Abrufen von MD5SUMS: $_" -Level 'ERROR'
    }
}

# Convert e.g. "7.1.6-166810" → "7.1.6.166810" for version comparison
$onlineVersionDotted = if ($installerFilename -match '\d+\.\d+\.\d+-\d+') { $Matches[0] -replace '-','.' } else { $onlineVersion }

# ── Installer ─────────────────────────────────────────────────────────────────
$localInstaller = Get-InstallerFilePath -Directory $InstallationFolder -Filter $installerFilter
$localInstallerVersion = $null
if ($localInstaller) {
    $localInstallerVersion = (Get-Item $localInstaller.FullName).VersionInfo.FileVersion
}

Write-Host ""
Write-Host "Lokale Version (Installer): $localInstallerVersion"  -ForegroundColor Cyan
Write-Host "Online Version:             $onlineVersionDotted"     -ForegroundColor Cyan
Write-Host ""

if ($onlineVersion -and $installerFilename) {
    $needDownload = $false
    if ($localInstallerVersion) {
        $needDownload = Compare-VersionComponents -Local $localInstallerVersion -Remote $onlineVersionDotted
    } else {
        $needDownload = $true
        Write-Host "$ProgramName nicht gefunden. Herunterladen..." -ForegroundColor DarkGray
    }

    if ($needDownload) {
        $destPath = Join-Path $InstallationFolder $installerFilename
        $ok       = Invoke-DownloadFile -Url "https://download.virtualbox.org/virtualbox/$onlineVersion/$installerFilename" -OutFile $destPath
        if ($ok -and (Test-Path $destPath)) {
            if ($localInstaller) { Remove-PathSafe -Path $localInstaller.FullName | Out-Null }
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "Installer aktualisiert: $destPath" -Level 'SUCCESS'
        } else {
            Remove-Item -Path $destPath -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update für Installer." -Level 'INFO'
    }
}

# ── Extension Pack ─────────────────────────────────────────────────────────────
$localExtPack = Get-InstallerFilePath -Directory $InstallationFolder -Filter $extPackFilter
$localExtPackVersion = $null
if ($localExtPack) {
    $epVer = [regex]::Match($localExtPack.Name, '(\d+\.\d+\.\d+(?:[-.]\d+)?)').Value
    $localExtPackVersion = $epVer -replace '-', '.'
}

Write-Host ""
Write-Host "Lokale Version (Extension Pack): $localExtPackVersion"  -ForegroundColor Cyan
Write-Host "Online Version:                  $onlineVersionDotted"  -ForegroundColor Cyan
Write-Host ""

if ($onlineVersion -and $extPackFilename) {
    $needExtPack = $false
    if ($localExtPackVersion) {
        $needExtPack = Compare-VersionComponents -Local $localExtPackVersion -Remote $onlineVersionDotted
    } else {
        $needExtPack = $true
        Write-Host "$ProgramName Extension Pack nicht gefunden. Herunterladen..." -ForegroundColor DarkGray
    }

    if ($needExtPack) {
        $destExt = Join-Path $InstallationFolder $extPackFilename
        $ok = Invoke-DownloadFile -Url "https://download.virtualbox.org/virtualbox/$onlineVersion/$extPackFilename" -OutFile $destExt
        if ($ok -and (Test-Path $destExt)) {
            if ($localExtPack) { Remove-PathSafe -Path $localExtPack.FullName | Out-Null }
            Write-Host "$ProgramName Extension Pack wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "Extension Pack aktualisiert: $destExt" -Level 'SUCCESS'
        } else {
            Remove-Item -Path $destExt -Force -ErrorAction SilentlyContinue
            Write-Host "Download ist fehlgeschlagen. Extension Pack wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Extension Pack Download fehlgeschlagen." -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. Extension Pack ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update für Extension Pack." -Level 'INFO'
    }
}

Write-Host ""

# ── Installed VirtualBox (exe) ─────────────────────────────────────────────────
$vbExe = "C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"
$installedVersion = $null
if (Test-Path $vbExe) {
    $installedVersion = (Get-Item $vbExe).VersionInfo.ProductVersion
}

# Re-read local installer version
$localInstaller = Get-InstallerFilePath -Directory $InstallationFolder -Filter $installerFilter
$localVersion   = if ($localInstaller) { (Get-Item $localInstaller.FullName).VersionInfo.FileVersion } else { $null }

$InstallVirtBox = $false
$ProgramName    = "Oracle VM VirtualBox"
if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'
    try { $InstallVirtBox = [version]$installedVersion -lt [version]$localVersion } catch { $InstallVirtBox = $false }
    if ($InstallVirtBox) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "$ProgramName nicht installiert." -Level 'INFO'
}

# ── Installed Extension Pack (ExtPack.xml) ─────────────────────────────────────
$ProgramName     = "VirtualBox Extension Pack"
$epInstalledVer  = $null
$epLocalVer      = $null
$localExtPack2   = Get-InstallerFilePath -Directory $InstallationFolder -Filter $extPackFilter
if ($localExtPack2) {
    $ev = [regex]::Match($localExtPack2.Name, '(\d+\.\d+\.\d+(?:[-.]\d+)?)').Value
    $epLocalVer = $ev -replace '-', '.'
}

$epXml = $null
$epDir = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) "Oracle\VirtualBox\ExtensionPacks"
if (Test-Path $epDir) {
    $epFolder = Get-ChildItem -Path $epDir -Directory | Where-Object { $_.Name -like '*VirtualBox_Extension_Pack*' } | Select-Object -First 1
    if ($epFolder) { $epXml = Join-Path $epFolder.FullName "ExtPack.xml" }
}

if ($epXml -and (Test-Path $epXml)) {
    try {
        $xml           = [xml](Get-Content $epXml)
        $ver           = $xml.VirtualBoxExtensionPack.Version
        $epInstalledVer = "$($ver.'#text').$($ver.revision)"
    } catch { Write-DeployLog -Message "Fehler beim Lesen von ExtPack.xml: $_" -Level 'ERROR' }
}

$InstallVirtBoxExPack = $false
if ($epInstalledVer) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $epInstalledVer" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $epLocalVer"     -ForegroundColor Cyan
    Write-DeployLog -Message "EP Installiert: $epInstalledVer | Lokal: $epLocalVer" -Level 'INFO'
    try { $InstallVirtBoxExPack = [version]$epInstalledVer -lt [version]$epLocalVer } catch { $InstallVirtBoxExPack = $false }
    if ($InstallVirtBoxExPack) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "$ProgramName nicht installiert." -Level 'INFO'
}

Write-Host ""

# ── Install if needed ──────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
}
if ($InstallVirtBox) {
    & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $installScript -InstallVirtBox
}
if ($InstallVirtBoxExPack) {
    & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $installScript -InstallVirtBoxExPack
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
