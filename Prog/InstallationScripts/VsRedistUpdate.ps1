param(
    [switch]$InstallationFlagX86 = $false,
    [switch]$InstallationFlagX64 = $false
)

$ProgramName = "VC Redist"
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

$localFilePathX86 = Join-Path $InstallationFolder "AutoIt_Scripts\VC_redist.x86.exe"
$localFilePathX64 = Join-Path $InstallationFolder "VirtualBox\VC_redist.x64.exe"

# ── Local versions ─────────────────────────────────────────────────────────────
function Get-FileProductVersion ([string]$Path) {
    if ($Path -and (Test-Path $Path)) {
        $v = (Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
        if ($v) { return $v }
    }
    return "0.0.0"
}

$localVersionX86 = Get-FileProductVersion -Path $localFilePathX86
$localVersionX64 = Get-FileProductVersion -Path $localFilePathX64
Write-DeployLog -Message "Lokale Versionen: X86=$localVersionX86 | X64=$localVersionX64" -Level 'DEBUG'

# ── Online version ─────────────────────────────────────────────────────────────
$onlineVersion   = $null
$downloadLinkX64 = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
$downloadLinkX86 = "https://aka.ms/vs/17/release/vc_redist.x86.exe"

$headers = @{ 'User-Agent'='InstallationScripts/1.0'; 'Accept'='text/plain' }
if ($GitHubToken) { $headers['Authorization'] = "token $GitHubToken" }

try {
    # Primary: GitHub Gist community tracker
    $gist = Invoke-RestMethod -Uri 'https://gist.githubusercontent.com/ChuckMichael/7366c38f27e524add3c54f710678c98b/raw/377c255f48319891068d29d6e4588ef5bc378a4e/vcredistr.md' -Headers $headers -ErrorAction Stop
    if ($gist -match '\((\d+\.\d+\.\d+)\)') {
        $onlineVersion = $Matches[1]
        if ($gist -match 'https://aka\.ms/vc\d+/vc_redist\.x64\.exe') { $downloadLinkX64 = $Matches[0] }
        if ($gist -match 'https://aka\.ms/vc\d+/vc_redist\.x86\.exe') { $downloadLinkX86 = $Matches[0] }
    }
    Write-DeployLog -Message "Online-Version (Gist): $onlineVersion" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Gist-Abfrage fehlgeschlagen: $_ — Fallback Filehorse" -Level 'WARNING'
    try {
        $fc = (Invoke-WebRequest -Uri 'https://www.filehorse.com/download-microsoft-visual-c-redistributable-package-64/' -UseBasicParsing -ErrorAction Stop).Content
        if ($fc -match '14\.\d+\.\d+\.\d+') { $onlineVersion = $Matches[0] }
        Write-DeployLog -Message "Online-Version (Filehorse): $onlineVersion" -Level 'INFO'
    } catch {
        Write-DeployLog -Message "Filehorse-Abfrage ebenfalls fehlgeschlagen: $_" -Level 'ERROR'
    }
}

Write-Host ""
Write-Host "Lokale Version X86: $localVersionX86" -ForegroundColor Cyan
Write-Host "Lokale Version X64: $localVersionX64" -ForegroundColor Cyan
Write-Host "Online Version:     $onlineVersion"   -ForegroundColor Cyan
Write-Host ""

# ── Helper: version less-than, tolerant of extra text ─────────────────────────
function Test-VersionLess ([string]$a, [string]$b) {
    try {
        $va = [version]($a -replace '[^\d\.]','')
        $vb = [version]($b -replace '[^\d\.]','')
        return $va -lt $vb
    } catch { return $true }
}

# ── Download X86 if needed ─────────────────────────────────────────────────────
if ($onlineVersion -and (Test-VersionLess $localVersionX86 $onlineVersion)) {
    $tempX86 = Join-Path $env:TEMP "VC_redist.x86.exe"
    Write-DeployLog -Message "Download X86: $downloadLinkX86" -Level 'INFO'
    $ok = Invoke-DownloadFile -Url $downloadLinkX86 -OutFile $tempX86
    if ($ok -and (Test-Path $tempX86)) {
        $destDir = Split-Path $localFilePathX86 -Parent
        if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
        try { Remove-Item -Path $localFilePathX86 -Force -ErrorAction SilentlyContinue } catch {}
        Move-Item -Path $tempX86 -Destination $localFilePathX86 -Force
        Write-Host "VC Redist x86 wurde aktualisiert." -ForegroundColor Green
        Write-DeployLog -Message "X86 aktualisiert: $localFilePathX86" -Level 'SUCCESS'
    } else {
        Remove-Item -Path $tempX86 -Force -ErrorAction SilentlyContinue
        Write-DeployLog -Message "Download X86 fehlgeschlagen." -Level 'ERROR'
    }
} else {
    Write-DeployLog -Message "Kein Update für X86 erforderlich." -Level 'DEBUG'
}

# ── Download X64 if needed ─────────────────────────────────────────────────────
if ($onlineVersion -and (Test-VersionLess $localVersionX64 $onlineVersion)) {
    $tempX64 = Join-Path $env:TEMP "VC_redist.x64.exe"
    Write-DeployLog -Message "Download X64: $downloadLinkX64" -Level 'INFO'
    $ok = Invoke-DownloadFile -Url $downloadLinkX64 -OutFile $tempX64
    if ($ok -and (Test-Path $tempX64)) {
        $destDir64 = Split-Path $localFilePathX64 -Parent
        if (-not (Test-Path $destDir64)) { New-Item -Path $destDir64 -ItemType Directory -Force | Out-Null }
        try { Remove-Item -Path $localFilePathX64 -Force -ErrorAction SilentlyContinue } catch {}
        Move-Item -Path $tempX64 -Destination $localFilePathX64 -Force
        Write-Host "VC Redist x64 wurde aktualisiert." -ForegroundColor Green
        Write-DeployLog -Message "X64 aktualisiert: $localFilePathX64" -Level 'SUCCESS'
    } else {
        Remove-Item -Path $tempX64 -Force -ErrorAction SilentlyContinue
        Write-DeployLog -Message "Download X64 fehlgeschlagen." -Level 'ERROR'
    }
} else {
    Write-DeployLog -Message "Kein Update für X64 erforderlich." -Level 'DEBUG'
}

# ── Re-read local versions ─────────────────────────────────────────────────────
$localVersionX86 = Get-FileProductVersion -Path $localFilePathX86
$localVersionX64 = Get-FileProductVersion -Path $localFilePathX64

# ── Installed X86 ─────────────────────────────────────────────────────────────
$installedInfoX86 = Get-RegistryVersion -DisplayNameLike '*-2022 Redistributable (x86)*'
$installedX86     = if ($installedInfoX86) { $installedInfoX86.VersionRaw } else { $null }
$Installx86       = $false

$ProgramName = "VC Redist x86"
if ($installedX86) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedX86"    -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersionX86" -ForegroundColor Cyan
    Write-DeployLog -Message "X86 Installiert: $installedX86 | Lokal: $localVersionX86" -Level 'INFO'
    $Installx86 = Test-VersionLess $installedX86 $localVersionX86
    if ($Installx86) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "X86 nicht in Registry." -Level 'INFO'
}

# ── Installed X64 ─────────────────────────────────────────────────────────────
$installedInfoX64 = Get-RegistryVersion -DisplayNameLike '*-2022 Redistributable (x64)*'
$installedX64     = if ($installedInfoX64) { $installedInfoX64.VersionRaw } else { $null }
$Installx64       = $false

$ProgramName = "VC Redist x64"
if ($installedX64) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedX64"    -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersionX64" -ForegroundColor Cyan
    Write-DeployLog -Message "X64 Installiert: $installedX64 | Lokal: $localVersionX64" -Level 'INFO'
    $Installx64 = Test-VersionLess $installedX64 $localVersionX64
    if ($Installx64) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "X64 nicht in Registry." -Level 'INFO'
}

Write-Host ""

# ── Install if needed ──────────────────────────────────────────────────────────
if ($Installx86 -or $InstallationFlagX86) {
    Write-Host "Microsoft Visual C++ x86 wird installiert" -ForegroundColor Magenta
    Write-DeployLog -Message "Starte Installation VC Redist x86." -Level 'INFO'
    $vcX86 = Get-ChildItem -Path (Join-Path $Serverip "Daten\Prog\AutoIt_Scripts\VC_redist*.exe") -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($vcX86) {
        & $vcX86.FullName /install /passive /qn /norestart | Out-Null
        Write-DeployLog -Message "VC x86 Installation gestartet: $($vcX86.FullName)" -Level 'SUCCESS'
    } else {
        Write-DeployLog -Message "Kein VC x86 Installer auf Server gefunden." -Level 'ERROR'
    }
}

if ($Installx64 -or $InstallationFlagX64) {
    Write-Host "Microsoft Visual C++ x64 wird installiert" -ForegroundColor Magenta
    Write-DeployLog -Message "Starte Installation VC Redist x64." -Level 'INFO'
    $vcX64 = Get-ChildItem -Path (Join-Path $Serverip "Daten\Prog\VirtualBox\VC*.exe") -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($vcX64) {
        & $vcX64.FullName /install /passive /qn /norestart | Out-Null
        Write-DeployLog -Message "VC x64 Installation gestartet: $($vcX64.FullName)" -Level 'SUCCESS'
    } else {
        Write-DeployLog -Message "Kein VC x64 Installer auf Server gefunden." -Level 'ERROR'
    }
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
