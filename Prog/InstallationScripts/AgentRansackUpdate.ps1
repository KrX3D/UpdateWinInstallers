param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Agent Ransack"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$onlineVersionUrl = "https://www.mythicsoft.com/agentransack/"
$installScript    = "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1"

# ── Helper: extract build number from filename ────────────────────────────────
function Get-AgentRansackBuildNumber {
    param([string]$FileName)
    # Matches: agentransack_3555.exe  agentransack_x64_msi_3555.zip  agentransack_x64_msi_3555.msi
    $m = [regex]::Match([System.IO.Path]::GetFileNameWithoutExtension($FileName), '(\d{4,})$')
    if ($m.Success) { return $m.Groups[1].Value }
    # fallback: any 4+ digit number in name
    $m = [regex]::Match($FileName, '(\d{4,})')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

# ── Local installer file (MSI preferred, EXE fallback) ────────────────────────
$localFileObj = Get-ChildItem -Path $InstallationFolder -Filter "agentransack*" -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '\.(exe|msi)$' } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

$localFilePath = if ($localFileObj) { $localFileObj.FullName } else { $null }
$localVersion  = if ($localFilePath) { Get-AgentRansackBuildNumber $localFileObj.Name } else { $null }

Write-DeployLog -Message "Lokale Datei: $localFilePath | Version: $localVersion" -Level 'DEBUG'

# ── Scrape download page – prefer MSI zip, fall back to EXE ──────────────────
$onlineVersion = $null
$downloadUrl   = $null
$downloadIsMsi = $false

try {
    $html = (Invoke-WebRequest -Uri $onlineVersionUrl -UseBasicParsing -ErrorAction Stop).Content

    # Scrape EXE link (the only href on the page) then derive MSI zip from same token.
    # EXE pattern: /flp/{build}/{token}/agentransack_{build}.exe
    $exeMatch = [regex]::Match($html,
        'href="((?:https?:)?//download\.mythicsoft\.com/flp/(\d+)/([^"]+)/agentransack_\d+\.exe)"')

    if ($exeMatch.Success) {
        $rawHref       = $exeMatch.Groups[1].Value
        $onlineVersion = $exeMatch.Groups[2].Value
        $token         = $exeMatch.Groups[3].Value
        $exeAbsolute   = if ($rawHref -match '^//') { "https:$rawHref" } else { $rawHref }
        $basePath      = $exeAbsolute -replace '[^/]+$', ''   # strip filename, keep trailing slash

        # Construct MSI zip URL with same token: agentransack_x64_msi_{build}.zip
        $msiUrl = "${basePath}agentransack_x64_msi_${onlineVersion}.zip"

        Write-DeployLog -Message "EXE-Link gescraped – Online-Version: $onlineVersion | Token: $token" -Level 'DEBUG'
        Write-DeployLog -Message "Pruefe MSI-Zip URL: $msiUrl" -Level 'DEBUG'

        # HEAD request to confirm MSI zip is available
        $msiAvailable = $false
        try {
            $head = Invoke-WebRequest -Uri $msiUrl -Method Head -UseBasicParsing -ErrorAction Stop
            $msiAvailable = ($head.StatusCode -eq 200)
        } catch { }

        if ($msiAvailable) {
            $downloadUrl   = $msiUrl
            $downloadIsMsi = $true
            Write-DeployLog -Message "MSI-Zip verfuegbar – wird bevorzugt: $downloadUrl" -Level 'INFO'
        } else {
            $downloadUrl   = $exeAbsolute
            $downloadIsMsi = $false
            Write-DeployLog -Message "MSI-Zip nicht erreichbar – Fallback auf EXE: $downloadUrl" -Level 'INFO'
        }
    } else {
        Write-DeployLog -Message "Kein Download-Link auf der Seite gefunden." -Level 'WARNING'
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Webseite: $_" -Level 'ERROR'
}

# ── Compare and download ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

if ($onlineVersion -and $downloadUrl) {
    $needsDownload = (-not $localVersion) -or ([int]$onlineVersion -gt [int]$localVersion)

    if ($needsDownload) {
        $fileName = Split-Path $downloadUrl -Leaf
        $tempFile = Join-Path $env:TEMP $fileName

        Write-DeployLog -Message "Lade herunter: $downloadUrl -> $tempFile" -Level 'INFO'
        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempFile

        if ($ok -and (Test-Path $tempFile)) {
            if ($downloadIsMsi) {
                # Extract zip, the MSI lands in $InstallationFolder
                Write-Host "Entpacke MSI-Paket..." -ForegroundColor Cyan
                Write-DeployLog -Message "Entpacke ZIP: $tempFile -> $InstallationFolder" -Level 'INFO'
                try {
                    Expand-Archive -Path $tempFile -DestinationPath $InstallationFolder -Force -ErrorAction Stop
                    Write-DeployLog -Message "ZIP erfolgreich entpackt." -Level 'SUCCESS'
                } catch {
                    Write-DeployLog -Message "Fehler beim Entpacken: $_" -Level 'ERROR'
                    Write-Host "Entpacken fehlgeschlagen!" -ForegroundColor Red
                }
                Remove-PathSafe -Path $tempFile | Out-Null
            } else {
                # EXE: copy directly into InstallationFolder
                $destPath = Join-Path $InstallationFolder $fileName
                Copy-FileSafe -Source $tempFile -Destination $destPath | Out-Null
                Remove-PathSafe -Path $tempFile | Out-Null
                Write-DeployLog -Message "EXE kopiert nach: $destPath" -Level 'SUCCESS'
            }

            # Remove old installer file (different version)
            if ($localFilePath -and (Test-Path $localFilePath)) {
                $oldBuild = Get-AgentRansackBuildNumber (Split-Path $localFilePath -Leaf)
                if ($oldBuild -ne $onlineVersion) {
                    Remove-PathSafe -Path $localFilePath | Out-Null
                    Write-DeployLog -Message "Alte Installationsdatei entfernt: $localFilePath" -Level 'INFO'
                }
            }

            Write-Host "$ProgramName wurde aktualisiert auf Version $onlineVersion." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName Installer aktualisiert auf $onlineVersion." -Level 'SUCCESS'
        } else {
            Write-Host "Download fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen: $downloadUrl" -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online-Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich (lokal=$localVersion, online=$onlineVersion)." -Level 'INFO'
    }
} elseif (-not $onlineVersion) {
    Write-Host "Online-Version konnte nicht ermittelt werden." -ForegroundColor Red
    Write-DeployLog -Message "Online-Version nicht abrufbar – Abbruch der Download-Prüfung." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file after potential download ───────────────────────────
$localFileObj  = Get-ChildItem -Path $InstallationFolder -Filter "agentransack*" -ErrorAction SilentlyContinue |
                 Where-Object { $_.Extension -match '\.(exe|msi)$' } |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 1
$localFilePath = if ($localFileObj) { $localFileObj.FullName } else { $null }
$localVersion  = if ($localFilePath) { Get-AgentRansackBuildNumber $localFileObj.Name } else { $null }

# ── Installed version (build = 3rd segment of DisplayVersion "2.1.3555") ─────
$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$Install       = $false

if ($installedInfo -and $installedInfo.VersionRaw) {
    $parts            = $installedInfo.VersionRaw -split '\.'
    $installedVersion = if ($parts.Count -ge 3) { $parts[2] } else { $installedInfo.VersionRaw }

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    if ($localVersion -and [int]$installedVersion -lt [int]$localVersion) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
        $Install = $true
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "$ProgramName nicht in Registry gefunden." -Level 'INFO'
}

Write-Host ""

# ── Trigger install script if needed ─────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -InstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
