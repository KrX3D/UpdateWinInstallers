param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Putty"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

# Source file (network share) and install destination
$sourcePath         = Join-Path $InstallationFolder "putty.exe"
$installDir         = "C:\Program Files (x86)\PuTTY"
$installPath        = Join-Path $installDir "putty.exe"
$desktopShortcut    = Join-Path ([Environment]::GetFolderPath("Desktop")) "PuTTY.lnk"

# ── Local (source) version ─────────────────────────────────────────────────────
$localVersion = $null
if (Test-Path $sourcePath) {
    $vi           = (Get-Item $sourcePath).VersionInfo
    $localVersion = "$($vi.ProductMajorPart).$($vi.ProductMinorPart)"
    Write-DeployLog -Message "Lokale Datei: putty.exe | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale putty.exe auf dem Share gefunden: $sourcePath" -Level 'WARNING'
}

# ── Online version ─────────────────────────────────────────────────────────────
$onlineVersion = $null
$downloadUrl   = $null

try {
    $html = (Invoke-WebRequest -Uri 'https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html' -UseBasicParsing -ErrorAction Stop).Content

    $vm = [regex]::Match($html, 'latest release \(([\d.]+)\)')
    if ($vm.Success) { $onlineVersion = $vm.Groups[1].Value }

    $lm = [regex]::Match($html, '<span class="downloadname">64-bit x86:</span>\s*<span class="downloadfile"><a href="(.*?putty\.exe)">')
    if ($lm.Success) { $downloadUrl = $lm.Groups[1].Value }

    Write-DeployLog -Message "Online-Version: $onlineVersion | URL: $downloadUrl" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der PuTTY-Seite: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer (updates the source file on the share) ──────────────────
if ($onlineVersion -and $downloadUrl -and $localVersion) {
    $isNewer = $false
    try { $isNewer = [version]$onlineVersion -gt [version]$localVersion } catch { $isNewer = $onlineVersion -ne $localVersion }

    if ($isNewer) {
        $tempPath = Join-Path $env:TEMP "putty.exe"
        $ok = Invoke-DownloadFile -Url $downloadUrl -OutFile $tempPath
        if ($ok -and (Test-Path $tempPath)) {
            try {
                if (Test-Path $sourcePath) { Remove-Item -Path $sourcePath -Force }
                Move-Item -Path $tempPath -Destination $sourcePath -Force
                Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                Write-DeployLog -Message "$ProgramName aktualisiert: $sourcePath" -Level 'SUCCESS'

                # Re-read local version
                $vi           = (Get-Item $sourcePath).VersionInfo
                $localVersion = "$($vi.ProductMajorPart).$($vi.ProductMinorPart)"
            } catch {
                Write-DeployLog -Message "Fehler beim Ersetzen der Quelldatei: $_" -Level 'ERROR'
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
                Write-Host "Fehler beim Aktualisieren. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
            }
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

# ── Installed (C:\Program Files\PuTTY) vs. source ─────────────────────────────
$installedVersion = $null
if (Test-Path $installPath) {
    $vi2              = (Get-Item $installPath).VersionInfo
    $installedVersion = "$($vi2.ProductMajorPart).$($vi2.ProductMinorPart).$($vi2.ProductBuildPart).$($vi2.ProductPrivatePart)"
}

$Install = $false
if ($installedVersion) {
    # Normalize source version to 4-part for fair comparison
    $vi3     = (Get-Item $sourcePath -ErrorAction SilentlyContinue)?.VersionInfo
    $srcVer4 = if ($vi3) { "$($vi3.ProductMajorPart).$($vi3.ProductMinorPart).$($vi3.ProductBuildPart).$($vi3.ProductPrivatePart)" } else { "0.0.0.0" }

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $srcVer4"          -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $srcVer4" -Level 'INFO'

    try { $Install = [version]$installedVersion -lt [version]$srcVer4 } catch { $Install = $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
    Write-DeployLog -Message "$ProgramName nicht in Installationspfad gefunden." -Level 'INFO'
}

Write-Host ""

# ── Copy / install if needed ───────────────────────────────────────────────────
if ($Install -or $InstallationFlag) {
    Write-Host "Putty wird kopiert" -ForegroundColor Cyan
    Write-DeployLog -Message "Starte Kopiervorgang: $sourcePath -> $installPath" -Level 'INFO'

    if (-not (Test-Path $installDir)) {
        New-Item -Path $installDir -ItemType Directory -Force | Out-Null
        Write-DeployLog -Message "Installationsverzeichnis erstellt: $installDir" -Level 'INFO'
    }

    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $installPath -Force
        Write-DeployLog -Message "Kopiert: $installPath" -Level 'SUCCESS'
    } else {
        Write-DeployLog -Message "Quelldatei nicht gefunden: $sourcePath" -Level 'ERROR'
    }

    # ── Desktop shortcut ──────────────────────────────────────────────────────
    try {
        $shell    = New-Object -ComObject WScript.Shell
        $lnk      = $shell.CreateShortcut($desktopShortcut)
        $lnk.TargetPath       = $installPath
        $lnk.WorkingDirectory = $installDir
        if (Test-Path $installPath) { $lnk.IconLocation = $installPath }
        $lnk.Save()
        Write-DeployLog -Message "Desktop-Verknüpfung erstellt: $desktopShortcut" -Level 'SUCCESS'
    } catch {
        Write-DeployLog -Message "Fehler beim Erstellen der Verknüpfung: $_" -Level 'ERROR'
    }

    # ── SSH Host Keys ──────────────────────────────────────────────────────────
    Write-Host "    SSH Host Keys werden gesetzt" -ForegroundColor Cyan
    $scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
    try {
        $PCName = & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
        Write-DeployLog -Message "PC-Name ermittelt: $PCName" -Level 'DEBUG'

        $backupFile = "$Serverip\Daten\Windows_Backup\$PCName\PuTTY_SSHHostKeys_Backup.reg"
        if (Test-Path $backupFile) {
            reg import $backupFile
            Write-DeployLog -Message "SSH Host Keys importiert: $backupFile" -Level 'SUCCESS'
        } else {
            Write-DeployLog -Message "SSH Host Key Backup nicht gefunden: $backupFile" -Level 'WARNING'
        }
    } catch {
        Write-DeployLog -Message "Fehler beim Abrufen des PC-Namens: $_" -Level 'ERROR'
    }

    # ── PuTTY Default Settings ─────────────────────────────────────────────────
    Write-Host "    PuTTY Default Settings werden gesetzt" -ForegroundColor Cyan
    $regPath = "HKCU:\Software\SimonTatham\PuTTY\Sessions\Default%20Settings"
    if (-not (Test-Path $regPath)) {
        New-Item -Path "HKCU:\Software\SimonTatham\PuTTY\Sessions" -Name "Default%20Settings" -Force | Out-Null
    }
    try {
        Set-ItemProperty -Path $regPath -Name "Font"        -Value "Terminal"
        Set-ItemProperty -Path $regPath -Name "FontCharSet" -Value 0x000000FF -Type DWord
        Write-DeployLog -Message "PuTTY Default Settings gesetzt (Font, FontCharSet)." -Level 'SUCCESS'
    } catch {
        Write-DeployLog -Message "Fehler beim Setzen der PuTTY Settings: $_" -Level 'ERROR'
    }
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
