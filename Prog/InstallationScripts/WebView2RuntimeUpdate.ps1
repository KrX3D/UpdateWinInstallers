param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Microsoft Edge Webview 2"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\WebView2RuntimeInstall.ps1"
$InstallFolder = "$InstallationFolder\ImageGlass"
$fileWildcard  = "MicrosoftEdgeWebview2Setup.exe"
$destFilePath  = Join-Path $InstallFolder $fileWildcard

# Local version (from registry)
$registryPaths = @(
    "HKCU:\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
)

$versions = @()
foreach ($regPath in $registryPaths) {
    if (Test-Path $regPath) {
        $v = (Get-ItemProperty -Path $regPath -Name 'pv' -ErrorAction SilentlyContinue).'pv'
        if ($v) { $versions += $v }
    }
}

$localVersion = if ($versions.Count -gt 0) {
    $versions | Sort-Object { [Version]$_ } | Select-Object -Last 1
} else {
    "0.0.0.0"
}
Write-DeployLog -Message "Lokale Version (Registry): $localVersion" -Level 'INFO'

# Online version via Microsoft Edge WebView2 download page
$webVersion = $null
try {
    $page = (Invoke-WebRequest -Uri 'https://developer.microsoft.com/de-de/microsoft-edge/webview2?form=MA13LH#download' -UseBasicParsing -ErrorAction Stop).Content
    if ($page -match '"__NUXT_DATA__".*?>(\[.+?\])<') {
        $parsed = $Matches[1] | ConvertFrom-Json
        $links  = @()
        foreach ($item in $parsed) {
            if ($item -match 'Microsoft\.WebView2\.FixedVersionRuntime\.(\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5})\.x64\.cab') {
                $links += [PSCustomObject]@{ Version = $Matches[1] }
            }
        }
        if ($links.Count -gt 0) {
            $webVersion = ($links | Sort-Object { [Version]$_.Version } -Descending | Select-Object -First 1).Version
        }
    }
    Write-DeployLog -Message "Online-Version: $webVersion" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der WebView2-Seite: $_" -Level 'ERROR'
}

Write-Host "Lokale Version: $localVersion" -ForegroundColor Cyan
Write-Host "Online Version: $webVersion"   -ForegroundColor Cyan
Write-Host ""

# Download / update installer file
$Install = $false

if ($webVersion) {
    $needDownload = $false
    try { $needDownload = [version]$localVersion -lt [version]$webVersion } catch { $needDownload = $localVersion -ne $webVersion }

    if ($needDownload) {
        $downloadLink = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
        Write-DeployLog -Message "Update verfuegbar: $localVersion -> $webVersion" -Level 'INFO'

        if (Test-Path $destFilePath) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
            $Install = $true
        } else {
            $ok = Invoke-DownloadFile -Url $downloadLink -OutFile $destFilePath
            if ($ok -and (Test-Path $destFilePath)) {
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
                $Install = $true
                Write-DeployLog -Message "Download erfolgreich: $destFilePath" -Level 'SUCCESS'
            } else {
                Remove-Item -Path $destFilePath -Force -ErrorAction SilentlyContinue
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
                Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
            }
        }
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
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
