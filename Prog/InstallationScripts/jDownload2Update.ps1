param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "jDownloader 2"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$NetworkShareDaten  = $config.NetworkShareDaten

$installScript    = "$NetworkShareDaten\Prog\InstallationScripts\Installation\jDownload2Install.ps1"
$InstallationFile = "$InstallationFolder\JDownloader2*.exe"

# ── Local file ────────────────────────────────────────────────────────────────
$FoundFile          = Get-ChildItem -Path $InstallationFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
$localInstaller     = if ($FoundFile) { Join-Path $InstallationFolder $FoundFile.Name } else { $null }
$localLastWriteTime = $null

if ($FoundFile) {
    try { $localLastWriteTime = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).LastWriteTime }
    catch { Write-DeployLog -Message "Konnte LastWriteTime nicht bestimmen: $_" -Level 'WARNING' }
} else {
    Write-DeployLog -Message "Keine lokale JDownloader-Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online build date ─────────────────────────────────────────────────────────
$onlineLastWriteTime = $null
try {
    $webContent = (Invoke-WebRequest 'https://svn.jdownloader.org/build.php' -UseBasicParsing -ErrorAction Stop).Content
    $dateMatch  = [regex]::Match($webContent, '(\w{3} \w{3} \d{2} \d{2}:\d{2}:\d{2} (CEST|CET) \d{4})')
    if ($dateMatch.Success) {
        $raw     = $dateMatch.Groups[1].Value
        $formats = @("ddd MMM dd HH:mm:ss 'CEST' yyyy", "ddd MMM dd HH:mm:ss 'CET' yyyy", "ddd MMM dd HH:mm:ss yyyy")
        foreach ($fmt in $formats) {
            try { $onlineLastWriteTime = [DateTime]::ParseExact($raw, $fmt, [System.Globalization.CultureInfo]::InvariantCulture); break } catch { }
        }
        if (-not $onlineLastWriteTime) {
            $clean = $raw -replace '\s+(CEST|CET)\s+', ' '
            try { $onlineLastWriteTime = [DateTime]::ParseExact($clean, "ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture) } catch { }
        }
    }
    Write-DeployLog -Message "Online Build-Datum: $onlineLastWriteTime" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Build-Website: $_" -Level 'ERROR'
}

Write-Host ""
Write-DeployLog -Message "Lokale Timestamp-Version: $localLastWriteTime" -Level 'INFO'
Write-DeployLog -Message "Online Timestamp-Version: $onlineLastWriteTime" -Level 'INFO'
Write-Host "Lokale Version: $localLastWriteTime"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineLastWriteTime" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer (via MEGAcmd) ───────────────────────────────────────────
$updateAvailable = $localLastWriteTime -and $onlineLastWriteTime -and ($localLastWriteTime -lt $onlineLastWriteTime)

if ($updateAvailable) {
    try {
        $pageContent = (Invoke-WebRequest 'https://jdownloader.org/jdownloader2' -UseBasicParsing -ErrorAction Stop).Content
        $linkMatch   = [regex]::Match($pageContent, '<td class="col2"> <a id="windows0" href="([^"]+)" class="urlextern" target="_blank"')

        if ($linkMatch.Success) {
            $downloadLink    = $linkMatch.Groups[1].Value
            $destinationPath = Join-Path $env:LOCALAPPDATA "MEGAcmd"
            Copy-Item -Path "$InstallationFolder\InstallationScripts\MEGAcmd" -Destination $destinationPath -Recurse -Force -ErrorAction SilentlyContinue

            Write-DeployLog -Message "Starte MEGAcmd Download: $downloadLink" -Level 'INFO'
            try {
                $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "$(Join-Path $destinationPath 'MEGAclient.exe') get $downloadLink $env:TEMP" -NoNewWindow -PassThru -ErrorAction Stop
                $proc.WaitForExit()
            } catch {
                Write-DeployLog -Message "MEGAcmd fehlgeschlagen: $_" -Level 'ERROR'
            }

            Get-Process -Name "MEGAcmdServer" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            Remove-Item -Path $destinationPath -Recurse -Force -ErrorAction SilentlyContinue

            $tempFile = Get-ChildItem -Path "$env:TEMP\JDownloader2*.exe" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($tempFile) {
                try { Set-ItemProperty -Path $tempFile.FullName -Name LastWriteTime -Value $onlineLastWriteTime } catch { }
                if ($localInstaller) { Remove-Item -Path $localInstaller -Recurse -Force -ErrorAction SilentlyContinue }
                try {
                    Move-Item -Path $tempFile.FullName -Destination $InstallationFolder -Force -ErrorAction Stop
                    Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                    Write-DeployLog -Message "$ProgramName aktualisiert." -Level 'SUCCESS'
                } catch {
                    Write-DeployLog -Message "Fehler beim Verschieben der Datei: $_" -Level 'ERROR'
                }
            } else {
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
                Write-DeployLog -Message "Temporäre Datei nicht gefunden nach MEGAcmd." -Level 'ERROR'
            }
        } else {
            Write-DeployLog -Message "Downloadlink konnte nicht aus JDownloader-Seite extrahiert werden." -Level 'WARNING'
        }
    } catch {
        Write-DeployLog -Message "Fehler beim Abrufen der JDownloader-Website: $_" -Level 'ERROR'
    }
} elseif ($localLastWriteTime -and $onlineLastWriteTime) {
    Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$FoundFile          = Get-ChildItem -Path $InstallationFile -File -ErrorAction SilentlyContinue | Select-Object -First 1
$localInstaller     = if ($FoundFile) { Join-Path $InstallationFolder $FoundFile.Name } else { $null }
$localLastWriteTime = $null
if ($FoundFile) {
    try { $localLastWriteTime = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).LastWriteTime } catch { }
}

# ── Installed version (TimestampVersion registry or build.json fallback) ──────
# jDownloader uses a non-standard timestamp in registry rather than a semver,
# so Get-RegistryVersion is not used here.
$registryEntry    = Get-InstalledSoftware -DisplayNameLike "$ProgramName*"
$installedVersion = $null

if ($registryEntry) {
    $raw = if ($registryEntry.PSObject.Properties['TimestampVersion']) { $registryEntry.TimestampVersion } else { $null }
    if ($raw) {
        $tz = if ((Get-Date).IsDaylightSavingTime()) { 'CEST' } else { 'CET' }
        try { $installedVersion = [DateTime]::ParseExact($raw, "ddd MMM dd HH:mm:ss '$tz' yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
        catch {
            try { $installedVersion = [DateTime]::ParseExact(($raw -replace '\s+(CEST|CET)\s+', ' '), "ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
            catch { Write-DeployLog -Message "Konnte TimestampVersion nicht parsen: $raw" -Level 'WARNING' }
        }
    } else {
        # Fallback: build.json
        $buildJson = "C:\Program Files\JDownloader\build.json"
        if (Test-Path $buildJson) {
            try {
                $ts = (Get-Content -Raw $buildJson | ConvertFrom-Json).buildDate
                if ($ts) {
                    $tz = if ((Get-Date).IsDaylightSavingTime()) { 'CEST' } else { 'CET' }
                    try { $installedVersion = [DateTime]::ParseExact($ts, "ddd MMM dd HH:mm:ss '$tz' yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
                    catch {
                        try { $installedVersion = [DateTime]::ParseExact(($ts -replace '\s+(CEST|CET)\s+', ' '), "ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture) }
                        catch { Write-DeployLog -Message "Konnte buildDate nicht parsen: $ts" -Level 'WARNING' }
                    }
                }
            } catch { Write-DeployLog -Message "Fehler beim Lesen von build.json: $_" -Level 'ERROR' }
        } else {
            Write-DeployLog -Message "build.json nicht gefunden: $buildJson" -Level 'WARNING'
        }
    }
}

$Install = $false
if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion"   -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localLastWriteTime" -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localLastWriteTime" -Level 'INFO'

    if ($localLastWriteTime -and ($installedVersion -ne $localLastWriteTime)) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        $Install = $true
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "$ProgramName nicht installiert." -Level 'INFO'
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
