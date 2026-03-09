param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Logitech G HUB"
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

$InstallationFolder = "$NetworkShareDaten\Treiber\LgHub"
$localFileFilter    = "lghub_installer*.exe"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder" -Level 'INFO'

# ── Local version (from filename) ─────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $localVersion = $localFile.Name -replace 'lghub_installer_|\.exe', ''
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden – überspringe Update-Check." -Level 'WARNING'
}

# ── Online version via Chrome headless ────────────────────────────────────────
$chromePath     = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$supportPageUrl = 'https://support.logi.com/hc/en-us/articles/360025298133'
$tempHtmlPath   = Join-Path ([System.IO.Path]::GetTempPath()) 'LGHUB.html'
$onlineVersion  = '0'
$downloadUrl    = 'https://download01.logi.com/web/ftp/pub/techsupport/gaming/lghub_installer.exe'

if ($localFile) {
    function Invoke-ChromeDump ([string]$ExtraArgs) {
        try {
            Start-Process -FilePath $chromePath `
                -ArgumentList "$ExtraArgs --dump-dom $supportPageUrl" `
                -RedirectStandardOutput $tempHtmlPath -NoNewWindow -Wait -ErrorAction Stop
        } catch {
            Write-DeployLog -Message "Chrome Dump fehlgeschlagen: $_" -Level 'ERROR'
        }
    }

    if (Test-Path $chromePath) {
        Invoke-ChromeDump '--headless=old --run-all-compositor-stages-before-draw --virtual-time-budget=60000'
        if ((Test-Path $tempHtmlPath) -and (Get-Item $tempHtmlPath).Length -eq 0) {
            Write-DeployLog -Message "Output leer; zweiter Versuch mit --headless." -Level 'WARNING'
            Invoke-ChromeDump '--headless --run-all-compositor-stages-before-draw --virtual-time-budget=60000'
        }

        if ((Test-Path $tempHtmlPath) -and (Get-Item $tempHtmlPath).Length -gt 0) {
            $html = Get-Content $tempHtmlPath -Raw

            $vm = [regex]::Match($html, '<b><span>Software Version: </span></b>(\d+\.\d+\.\d+)')
            if ($vm.Success) {
                $onlineVersion = $vm.Groups[1].Value
                Write-DeployLog -Message "Online-Version: $onlineVersion" -Level 'INFO'
            } else {
                Write-DeployLog -Message "Version nicht im HTML gefunden; Fallback wird verwendet." -Level 'WARNING'
            }

            $lm = [regex]::Match($html, '<a class="download-button" href="(.*?)" target="_blank">Download Now</a>')
            if ($lm.Success) {
                $downloadUrl = $lm.Groups[1].Value
                Write-DeployLog -Message "Download-URL: $downloadUrl" -Level 'INFO'
            }
        } else {
            Write-Host "Chrome konnte keine gültigen Daten abrufen. Verwende Fallback-Werte." -ForegroundColor Yellow
            Write-DeployLog -Message "Kein valides HTML nach Chrome-Aufruf." -Level 'WARNING'
        }
    } else {
        Write-Host "Chrome nicht gefunden. Verwende Fallback-Werte." -ForegroundColor Yellow
        Write-DeployLog -Message "Chrome nicht gefunden: $chromePath" -Level 'WARNING'
    }

    Write-Host ""
    Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
    Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
    Write-Host ""

    # ── Download if newer (integer component comparison) ──────────────────────
    if ($onlineVersion -ne '0' -and
        $localVersion  -match '^\d+\.\d+\.\d+$' -and
        $onlineVersion -match '^\d+\.\d+\.\d+$') {

        $lc = $localVersion  -split '\.' | ForEach-Object { [int]$_ }
        $oc = $onlineVersion -split '\.' | ForEach-Object { [int]$_ }
        $isNewer = $oc[0] -gt $lc[0] -or
                   ($oc[0] -eq $lc[0] -and $oc[1] -gt $lc[1]) -or
                   ($oc[0] -eq $lc[0] -and $oc[1] -eq $lc[1] -and $oc[2] -gt $lc[2])

        if ($isNewer) {
            $destPath = Join-Path $InstallationFolder "lghub_installer_$onlineVersion.exe"

            [void](Invoke-InstallerDownload `
                -Url                $downloadUrl `
                -OutFile            $destPath `
                -ConfirmDownload `
                -ReplaceOld `
                -RemoveFiles        @($localFile.FullName) `
                -KeepFiles          @($destPath) `
                -EmitHostStatus `
                -SuccessHostMessage "$ProgramName wurde aktualisiert.." `
                -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
                -SuccessLogMessage  "$ProgramName erfolgreich aktualisiert: $destPath" `
                -FailureLogMessage  "Download fehlgeschlagen: $destPath" `
                -Context            $ProgramName)
        } else {
            Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
            Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
        }
    } else {
        Write-DeployLog -Message "Ungültige Versionsdaten – Update-Prüfung übersprungen." -Level 'WARNING'
    }

    Remove-Item -Path $tempHtmlPath -Force -ErrorAction SilentlyContinue
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = if ($localFile) { $localFile.Name -replace 'lghub_installer_|\.exe', '' } else { $null }

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    if ($localVersion -match '^\d+\.\d+\.\d+$' -and $installedVersion -match '^\d+\.\d+\.\d+$') {
        $ic = $installedVersion -split '\.' | ForEach-Object { [int]$_ }
        $lc = $localVersion     -split '\.' | ForEach-Object { [int]$_ }
        $Install = $ic[0] -lt $lc[0] -or
                   ($ic[0] -eq $lc[0] -and $ic[1] -lt $lc[1]) -or
                   ($ic[0] -eq $lc[0] -and $ic[1] -eq $lc[1] -and $ic[2] -lt $lc[2])
    }

    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht in Registry gefunden." -Level 'INFO'
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
