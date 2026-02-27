param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinSCP"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\WinScpInstall.ps1"
$localFileFilter = "WinSCP-*.exe"

# Local version (ProductVersion from installer)
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null
if ($localFile) {
    $localVersion = (Get-Item $localFile.FullName).VersionInfo.ProductVersion.Trim()
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Kein WinSCP-Installer gefunden." -Level 'WARNING'
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
    return
}

# Online version - scrape WinSCP download page for filename containing version
$latestVersion = $null
$downloadLink  = $null

function Find-WinScpVersionInContent ([string]$Content) {
    $m = [regex]::Match($Content, 'WinSCP-([0-9]+(?:\.[0-9]+)+)-Setup\.exe', 'IgnoreCase')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return $null
}

try {
    $mainPage = Invoke-WebRequest -Uri 'https://winscp.net/eng/downloads.php' -UseBasicParsing -ErrorAction Stop

    # Try main page HTML first
    $latestVersion = Find-WinScpVersionInContent $mainPage.Content

    # If found, probe the redirect page to get the actual direct download link
    if ($latestVersion) {
        try {
            $redirectUrl  = "https://winscp.net/download/WinSCP-$latestVersion-Setup.exe"
            $redirectPage = Invoke-WebRequest -Uri $redirectUrl -UseBasicParsing -ErrorAction Stop

            # Find .exe URL in redirect page
            $exeMatch = [regex]::Match($redirectPage.Content, '(https?://[^"''<>\s]+?WinSCP[^"''<>\s]+?\.exe)', 'IgnoreCase')
            if ($exeMatch.Success) {
                $downloadLink = $exeMatch.Groups[1].Value
            } else {
                # href fallback
                $hrefMatch = [regex]::Match($redirectPage.Content, 'href\s*=\s*["'']([^"'']+?\.exe(?:\?[^"'']*)?)["'']', 'IgnoreCase')
                if ($hrefMatch.Success) {
                    $rel = $hrefMatch.Groups[1].Value
                    $downloadLink = if ($rel -match '^https?://') { $rel }
                                    elseif ($rel -match '^//') { "https:$rel" }
                                    else { "https://winscp.net$rel" }
                }
            }
        } catch {
            Write-DeployLog -Message "Fehler beim Abrufen der Redirect-Seite: $_" -Level 'WARNING'
        }
    }

    # Fallback: try download landing page
    if (-not $latestVersion) {
        $landing = (Invoke-WebRequest -Uri 'https://winscp.net/download/' -UseBasicParsing -ErrorAction SilentlyContinue).Content
        if ($landing) { $latestVersion = Find-WinScpVersionInContent $landing }
    }

    Write-DeployLog -Message "Online-Version: $latestVersion | Download-Link: $downloadLink" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der WinSCP-Seite: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $latestVersion" -ForegroundColor Cyan
Write-Host ""

# Helper: version a > version b, split-based fallback
function Test-VersionGreater ([string]$a, [string]$b) {
    if (-not $a -or -not $b) { return $false }
    try { return [version]$a -gt [version]$b } catch {}
    $as = $a.Split('.') | ForEach-Object { try { [int]$_ } catch { 0 } }
    $bs = $b.Split('.') | ForEach-Object { try { [int]$_ } catch { 0 } }
    $n  = [Math]::Max($as.Length, $bs.Length)
    for ($i = 0; $i -lt $n; $i++) {
        $ai = if ($i -lt $as.Length) { $as[$i] } else { 0 }
        $bi = if ($i -lt $bs.Length) { $bs[$i] } else { 0 }
        if ($ai -gt $bi) { return $true }
        if ($ai -lt $bi) { return $false }
    }
    return $false
}

# Download if newer
if ($latestVersion -and (Test-VersionGreater $latestVersion $localVersion)) {
    Write-DeployLog -Message "Update verfuegbar: $localVersion -> $latestVersion" -Level 'INFO'

    # Build direct download URL if redirect page didn't yield one
    if (-not $downloadLink) {
        $downloadLink = "https://winscp.net/download/WinSCP-$latestVersion-Setup.exe"
    }

    try { $filename = [System.IO.Path]::GetFileName(([uri]$downloadLink).AbsolutePath) } catch { $filename = "WinSCP-$latestVersion-Setup.exe" }
    $destPath = Join-Path $InstallationFolder $filename
    $tempPath = "$destPath.part"

    $ok = Invoke-DownloadFile -Url $downloadLink -OutFile $tempPath
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

Write-Host ""

# Re-evaluate local file and installed version
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = if ($localFile) { (Get-Item $localFile.FullName).VersionInfo.ProductVersion.Trim() } else { $null }

$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    try { $Install = [version]$installedVersion -lt [version]$localVersion } catch {
        $Install = Test-VersionGreater $localVersion $installedVersion
    }
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
