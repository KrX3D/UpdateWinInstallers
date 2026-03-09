param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "FiiO"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config     = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$Serverip   = $config.Serverip
$PSHostPath = $config.PSHostPath
# Dot-source config to expose $NetworkShareDaten and other extended variables
. (Get-SharedConfigPath -ScriptRoot $PSScriptRoot)

$FiiODriverFolder = "$NetworkShareDaten\Treiber\FiiO_Verstarker"
$localPattern     = "FiiO_v*.msi"
$installScript    = "$Serverip\Daten\Prog\InstallationScripts\Installation\FiiOInstallation.ps1"

Write-DeployLog -Message "FiiO Treiber-Ordner: $FiiODriverFolder" -Level 'INFO'

# ── Local version (from filename) ─────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $FiiODriverFolder -Filter $localPattern
$localVersion = $null

if ($localFile) {
    if ($localFile.Name -match 'FiiO_v([\d\.]+)') { $localVersion = $Matches[1] }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'INFO'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version ────────────────────────────────────────────────────────────
$pageUrl    = 'https://forum.fiio.com/note/showNoteContent.do?id=202105191527366657910&tid=17'
$verPattern = 'V([\d\.]+)\s+version\s+download\s+link:.*?<a[^>]*href="([^"]+)"[^>]*>.*?Click here.*?</a>.*?\(For Win 10/11'

$onlineVersion = $null
$downloadLink  = $null
$pageContent   = $null

try {
    $resp        = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -ErrorAction Stop
    $pageContent = $resp.Content
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Download-Seite: $_" -Level 'ERROR'
}

if ($pageContent) {
    $m = [regex]::Match($pageContent, $verPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($m.Success) {
        $onlineVersion = $m.Groups[1].Value
        $downloadLink  = $m.Groups[2].Value
        Write-DeployLog -Message "Online-Version: $onlineVersion | Link: $downloadLink" -Level 'INFO'
    } else {
        Write-DeployLog -Message "Keine passende Version auf der Webseite gefunden." -Level 'WARNING'
    }
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download and extract MSI if newer ────────────────────────────────────────
if ($localFile -and $onlineVersion -and $downloadLink) {
    $needUpdate = try { [version]$onlineVersion -gt [version]$localVersion } catch { $onlineVersion -ne $localVersion }

    if ($needUpdate) {
        Write-DeployLog -Message "Neuere Version verfügbar: $onlineVersion > $localVersion" -Level 'INFO'

        if ($downloadLink -notmatch '^https?://') {
            $downloadUrl = "https://fiio-firmware.fiio.net/DAC/FiiO%20USB%20DAC%20driver%20V$onlineVersion.exe"
        } else {
            $downloadUrl = $downloadLink
        }

        $tempFolder  = Join-Path $env:TEMP "FiiO_Update_$onlineVersion"
        $tempExePath = Join-Path $tempFolder "FiiO_USB_DAC_driver_V$onlineVersion.exe"
        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null

        $ok = $false
        try {
            $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempExePath
        } catch {
            Write-DeployLog -Message "Download-Fehler: $_" -Level 'ERROR'
        }

        if ($ok -and (Test-Path $tempExePath)) {
            try {
                $extractFolder = Join-Path $tempFolder 'Extracted'
                New-Item -Path $extractFolder -ItemType Directory -Force | Out-Null

                $sevenZip = 'C:\Program Files\7-Zip\7z.exe'
                if (Test-Path $sevenZip) {
                    & $sevenZip x "$tempExePath" "-o$extractFolder" -y 2>&1 | Out-Null
                } else {
                    Start-Process -FilePath $tempExePath -ArgumentList "/S /D=$extractFolder" -Wait -NoNewWindow
                }

                $extractedMsi = Get-ChildItem -Path (Join-Path $extractFolder 'x64') -Filter '*.msi' -ErrorAction SilentlyContinue | Select-Object -First 1

                if ($extractedMsi) {
                    $newMsiName = "FiiO_v${onlineVersion}.msi"
                    $newMsiPath = Join-Path $FiiODriverFolder $newMsiName
                    Copy-Item -Path $extractedMsi.FullName -Destination $newMsiPath -Force

                    if (Test-Path $newMsiPath) {
                        if ($localFile -and (Test-Path $localFile.FullName)) {
                            Remove-PathSafe -Path $localFile.FullName | Out-Null
                        }
                        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                        Write-DeployLog -Message "$ProgramName aktualisiert: $newMsiPath" -Level 'SUCCESS'
                        $localFile = Get-Item $newMsiPath
                    } else {
                        Write-DeployLog -Message "MSI konnte nicht kopiert werden." -Level 'ERROR'
                    }
                } else {
                    Write-DeployLog -Message "MSI-Datei nicht im x64-Ordner gefunden." -Level 'ERROR'
                }
            } catch {
                Write-DeployLog -Message "Extraktionsfehler: $_" -Level 'ERROR'
            } finally {
                Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "Download fehlgeschlagen: $tempExePath" -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFile) {
    Write-DeployLog -Message "Keine lokale Installationsdatei – Update-Vergleich übersprungen." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $FiiODriverFolder -Filter $localPattern
$localVersion = $null
if ($localFile -and $localFile.Name -match 'FiiO_v([\d\.]+)') { $localVersion = $Matches[1] }

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "*FiiO*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName Treiber ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    $Install = try { [version]$installedVersion -lt [version]$localVersion } catch { $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName Treiber installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "$ProgramName Treiber nicht in Registry gefunden." -Level 'INFO'
}

Write-Host ""

# ── Install if needed ─────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -InstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
