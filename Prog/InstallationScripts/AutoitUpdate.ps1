param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Autoit"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot
Write-DeployLog -Message "Script gestartet mit InstallationFlag: $InstallationFlag" -Level 'INFO'

$config            = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = "$($config.InstallationFolder)\AutoIt_Scripts"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$autoItDownloadUrl  = "https://www.autoitscript.com/site/autoit/downloads/"
$sciTEDownloadUrl   = "https://www.autoitscript.com/cgi-bin/getfile.pl?../autoit3/scite/download/SciTE4AutoIt3.exe"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1"

# ── Local file discovery ──────────────────────────────────────────────────────
$localAutoItFile = Get-ChildItem -Path "$InstallationFolder\autoit-v3-setup*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
$localSciTEFile  = Get-ChildItem -Path "$InstallationFolder\SciTE4AutoIt3*.exe"  -ErrorAction SilentlyContinue | Select-Object -Last 1

$localAutoItVersion = if ($localAutoItFile) {
    Get-InstallerFileVersion -FilePath $localAutoItFile.FullName -Source FileVersion
} else { $null }

$localSciTEVersion = if ($localSciTEFile) {
    Get-InstallerFileVersion -FilePath $localSciTEFile.FullName -Source ProductVersion
} else { $null }

Write-DeployLog -Message "Lokale AutoIt Datei:  $(if ($localAutoItFile) { $localAutoItFile.Name } else { 'None' }) | Version: $localAutoItVersion" -Level 'DEBUG'
Write-DeployLog -Message "Lokale SciTE Datei:   $(if ($localSciTEFile)  { $localSciTEFile.Name  } else { 'None' }) | Version: $localSciTEVersion"  -Level 'DEBUG'

# ── Online AutoIt version + download link ─────────────────────────────────────
# Version found in page content: v3.3.16.1 style
# Download link is a relative URL appended to the getfile.pl base
$autoItInfo = Get-OnlineInstallerLink `
    -Url           $autoItDownloadUrl `
    -LinkRegex     '(?<=href="\/cgi-bin\/getfile\.pl\?)([^"]+autoit-v3-setup[^"]*)' `
    -LinkPrefix    'https://www.autoitscript.com/cgi-bin/getfile.pl?' `
    -VersionRegex  'v(\d+\.\d+\.\d+\.\d+)' `
    -VersionSource Content `
    -Context       'AutoIt'

if ($autoItInfo.DownloadUrl -and $autoItInfo.Version) {
    $onlineAutoItVersion = $autoItInfo.Version
    $autoItDownloadLink  = $autoItInfo.DownloadUrl
    $filename            = Split-Path -Path $autoItDownloadLink -Leaf

    Write-Host ""
    Write-Host "$ProgramName Lokale Version: $localAutoItVersion" -ForegroundColor Cyan
    Write-Host "$ProgramName Online Version: $onlineAutoItVersion" -ForegroundColor Cyan
    Write-Host ""

    # String comparison preserved from original
    if ($onlineAutoItVersion -gt $localAutoItVersion) {
        $autoItSavePath = Join-Path $env:TEMP $filename

        $ok = Invoke-DownloadFile -Url $autoItDownloadLink -OutFile $autoItSavePath
        Write-DeployLog -Message "AutoIt Download: $(if ($ok) { 'OK' } else { 'FEHLGESCHLAGEN' })" -Level $(if ($ok) { 'SUCCESS' } else { 'ERROR' })

        if ($ok -and (Test-Path $autoItSavePath)) {
            # If somehow a zip arrives, extract it; exe files pass through as-is
            if ($autoItSavePath -match '\.zip$') {
                try {
                    Expand-Archive -Path $autoItSavePath -DestinationPath $env:TEMP -Force
                    Write-DeployLog -Message "AutoIt Archiv entpackt nach $env:TEMP" -Level 'SUCCESS'
                } catch {
                    Write-DeployLog -Message "Entpacken fehlgeschlagen: $($_.Exception.Message)" -Level 'ERROR'
                }
            }

            # Remove old local installer
            if ($localAutoItFile) { Remove-PathSafe -Path $localAutoItFile.FullName | Out-Null }

            # Locate the new exe (extracted or the downloaded exe itself)
            $newAutoItFile = (Get-ChildItem -Path "$env:TEMP\autoit*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1)?.FullName
            if (-not $newAutoItFile -and $autoItSavePath -match '\.exe$' -and (Test-Path $autoItSavePath)) {
                $newAutoItFile = $autoItSavePath
            }

            if ($newAutoItFile -and (Test-Path $newAutoItFile)) {
                Move-Item -Path $newAutoItFile -Destination $InstallationFolder -Force
                Write-DeployLog -Message "AutoIt Datei verschoben nach $InstallationFolder" -Level 'SUCCESS'
            } else {
                Write-DeployLog -Message "Neue AutoIt-Datei nicht gefunden zum Verschieben." -Level 'ERROR'
                Write-Host "Download ist fehlgeschlagen. $filename wurde nicht aktualisiert." -ForegroundColor Red
            }

            # Cleanup temp zip if applicable
            if ($autoItSavePath -match '\.zip$') { Remove-PathSafe -Path $autoItSavePath | Out-Null }

            # ── SciTE download (bundled with AutoIt update) ───────────────────
            $sciTESavePath = Join-Path $env:TEMP "SciTE4AutoIt3.exe"

            $sciOk = Invoke-DownloadFile -Url $sciTEDownloadUrl -OutFile $sciTESavePath
            Write-DeployLog -Message "SciTE Download: $(if ($sciOk) { 'OK' } else { 'FEHLGESCHLAGEN' })" -Level $(if ($sciOk) { 'SUCCESS' } else { 'ERROR' })

            if ($sciOk -and (Test-Path $sciTESavePath)) {
                if ($localSciTEFile) { Remove-PathSafe -Path $localSciTEFile.FullName | Out-Null }

                $newSciteFile = (Get-ChildItem -Path "$env:TEMP\SciTE4AutoIt3*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1)?.FullName
                if ($newSciteFile -and (Test-Path $newSciteFile)) {
                    Move-Item -Path $newSciteFile -Destination $InstallationFolder -Force
                    Write-DeployLog -Message "SciTE Datei verschoben nach $InstallationFolder" -Level 'SUCCESS'
                } else {
                    Write-DeployLog -Message "Neue SciTE-Datei nicht gefunden zum Verschieben." -Level 'ERROR'
                    Write-Host "Download ist fehlgeschlagen. SciTE4AutoIt3.exe wurde nicht aktualisiert." -ForegroundColor Red
                }
            } else {
                Write-Host "Download ist fehlgeschlagen. SciTE4AutoIt3.exe wurde nicht aktualisiert." -ForegroundColor Red
            }

            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert." -Level 'SUCCESS'
        } else {
            Write-Host "Download ist fehlgeschlagen. $filename wurde nicht aktualisiert." -ForegroundColor Red
            Write-DeployLog -Message "AutoIt Download fehlgeschlagen: $autoItSavePath" -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "Konnte AutoIt Version oder Download-Link nicht ermitteln." -Level 'WARNING'
}

Write-Host ""

# ── Re-read local versions after potential update ─────────────────────────────
$localAutoItFile = Get-ChildItem -Path "$InstallationFolder\autoit-v3-setup*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
$localSciTEFile  = Get-ChildItem -Path "$InstallationFolder\SciTE4AutoIt3*.exe"  -ErrorAction SilentlyContinue | Select-Object -Last 1

$localAutoItVersion = if ($localAutoItFile) {
    Get-InstallerFileVersion -FilePath $localAutoItFile.FullName -Source FileVersion
} else { $null }

$localSciTEVersion = if ($localSciTEFile) {
    Get-InstallerFileVersion -FilePath $localSciTEFile.FullName -Source ProductVersion
} else { $null }

# ── AutoIt: installed version check ──────────────────────────────────────────
$autoItInstallInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$AutoItInstall     = $false

if ($autoItInstallInfo) {
    $installedVersion = $autoItInstallInfo.VersionRaw
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localAutoItVersion" -ForegroundColor Cyan
    Write-DeployLog -Message "AutoIt installiert: $installedVersion | Lokal: $localAutoItVersion" -Level 'INFO'

    try {
        $AutoItInstall = [version]$installedVersion -lt [version]$localAutoItVersion
        if ($AutoItInstall) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        } else {
            Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        }
    } catch {
        Write-DeployLog -Message "Versionsvergleich fehlgeschlagen: $($_.Exception.Message)" -Level 'WARNING'
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht in Registry gefunden." -Level 'DEBUG'
}

Write-Host ""

# ── SciTE: installed version check ───────────────────────────────────────────
$sciteInstallInfo = Get-RegistryVersion -DisplayNameLike "Scite*"
$SciteInstall     = $false

if ($sciteInstallInfo) {
    $installedVersion = $sciteInstallInfo.VersionRaw
    Write-Host "Scite ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localSciTEVersion" -ForegroundColor Cyan
    Write-DeployLog -Message "Scite installiert: $installedVersion | Lokal: $localSciTEVersion" -Level 'INFO'

    try {
        $SciteInstall = [version]$installedVersion -lt [version]$localSciTEVersion
        if ($SciteInstall) {
            Write-Host "        Veraltete Scite ist installiert. Update wird gestartet." -ForegroundColor Magenta
        } else {
            Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        }
    } catch {
        Write-DeployLog -Message "Versionsvergleich fehlgeschlagen: $($_.Exception.Message)" -Level 'WARNING'
    }
} else {
    Write-DeployLog -Message "Scite nicht in Registry gefunden." -Level 'DEBUG'
}

Write-Host ""

# ── Install if needed ─────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag
}

if ($AutoItInstall) {
    & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $installScript -Autoit
    Write-DeployLog -Message "AutoIt Installationsskript aufgerufen." -Level 'DEBUG'
}

if ($SciteInstall) {
    & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $installScript -Scite
    Write-DeployLog -Message "SciTE Installationsskript aufgerufen." -Level 'DEBUG'
}

Write-Host ""
Write-DeployLog -Message "Script endet normal." -Level 'INFO'
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
