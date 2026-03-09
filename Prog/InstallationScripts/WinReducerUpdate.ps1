param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinReducerEX100"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
. (Get-SharedConfigPath -ScriptRoot $PSScriptRoot)   # exposes $NetworkShareDaten

$InstallationFolder  = "$NetworkShareDaten\Customize_Windows\Tools\WinReducerEX100"
$InstallationExe     = "$InstallationFolder\WinReducerEX100_x64.exe"
$destinationPath     = Join-Path $env:USERPROFILE "Desktop\WinReducerEX100"
$destinationFilePath = Join-Path $destinationPath "WinReducerEX100_x64.exe"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder | Exe: $InstallationExe" -Level 'DEBUG'

# Online version from WinReducer website
$webpageURL    = "https://www.winreducer.net/winreducer-ex-series.html"
$webpageContent = ""
try {
    $webpageContent = (Invoke-WebRequest -Uri $webpageURL -UseBasicParsing -ErrorAction Stop).Content
    Write-DeployLog -Message "Webseite abgerufen: $webpageURL" -Level 'DEBUG'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Webseite $webpageURL : $_" -Level 'ERROR'
}

$versionText = ""
if ($webpageContent) {
    $m = [regex]::Match($webpageContent, '(?<=<div class="paragraph" style="text-align:center;"><font size="6"><strong><font color="#fff">)v(.*?)(?=<\/font>)')
    $versionText = $m.Groups[1].Value.Trim()
}
Write-DeployLog -Message "Online-Version: $versionText" -Level 'INFO'

# Local version (FileVersion from exe)
$localVersion = ""
try {
    $localVersion = (Get-ItemProperty -Path $InstallationExe -ErrorAction Stop).VersionInfo.FileVersion
    Write-DeployLog -Message "Lokale Version: $localVersion" -Level 'DEBUG'
} catch {
    Write-DeployLog -Message "Fehler beim Lesen der lokalen Version: $_" -Level 'WARNING'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $versionText"   -ForegroundColor Cyan
Write-Host ""

# Download and update if newer
if ($versionText -ne '' -and $versionText -gt $localVersion) {
    Write-DeployLog -Message "Update verfuegbar: $localVersion -> $versionText" -Level 'INFO'

    # Extract download URL from page
    $dlMatch     = [regex]::Match($webpageContent, '<a class="wsite-button wsite-button-large wsite-button-highlight" href="(.+?)".*?>\s*<span class="wsite-button-inner">DOWNLOAD \(x64\)</span>')
    $partialUrl  = $dlMatch.Groups[1].Value
    $downloadURL = "https://www.winreducer.net$partialUrl"
    $downloadPath = Join-Path $env:TEMP "WinReducerEX100_x64_new.zip"

    Write-DeployLog -Message "Download-URL: $downloadURL" -Level 'DEBUG'

    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    try {
        Invoke-DownloadFile -Url $downloadURL -OutFile $downloadPath | Out-Null
        Write-DeployLog -Message "Download abgeschlossen: $downloadPath" -Level 'SUCCESS'
    } catch {
        Write-DeployLog -Message "Fehler beim Download: $_" -Level 'ERROR'
    } finally {
        $null = $wc.Dispose()
    }

    if (Test-Path $downloadPath) {
        try { Expand-Archive -Path $downloadPath -DestinationPath $env:TEMP -Force } catch {
            Write-DeployLog -Message "Fehler beim Entpacken: $_" -Level 'ERROR'
        }

        # Backup licence/config before replacing
        $backupDest = Join-Path $env:TEMP "WinReducer_EX_Series_x64\WinReducerEX100\HOME\SOFTWARE"
        try { Move-Item -Path (Join-Path $InstallationFolder "HOME\SOFTWARE\x64")              -Destination $backupDest -Force } catch {}
        try { Move-Item -Path (Join-Path $InstallationFolder "HOME\SOFTWARE\WinReducerEX100.xml") -Destination $backupDest -Force } catch {}

        # Replace installation folder
        try { Remove-Item -Path $InstallationFolder -Recurse -Force } catch {
            Write-DeployLog -Message "Fehler beim Entfernen des alten Installationsordners: $_" -Level 'WARNING'
        }
        $extractedFolder = Join-Path $env:TEMP "WinReducer_EX_Series_x64\WinReducerEX100"
        try {
            Move-Item -Path $extractedFolder -Destination (Split-Path $InstallationFolder -Parent) -Force
            Write-DeployLog -Message "Extrahierter Ordner verschoben nach: $(Split-Path $InstallationFolder -Parent)" -Level 'SUCCESS'
        } catch {
            Write-DeployLog -Message "Fehler beim Verschieben: $_" -Level 'ERROR'
        }

        # Cleanup temp
        Remove-Item -Path (Join-Path $env:TEMP "WinReducer_EX_Series_x64") -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $downloadPath -Force -ErrorAction SilentlyContinue

        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
        Write-DeployLog -Message "$ProgramName aktualisiert." -Level 'SUCCESS'
    } else {
        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
        Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
    }
} else {
    Write-Host "Kein Online Update verfuegbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
}

Write-Host ""

# Re-read local version after potential update
try {
    $localVersion = (Get-ItemProperty -Path $InstallationExe -ErrorAction Stop).VersionInfo.FileVersion
} catch {
    $localVersion = ""
}

# Installed version = file on Desktop
$installedVersion = $null
if (Test-Path $destinationFilePath) {
    try { $installedVersion = (Get-ItemProperty -Path $destinationFilePath).VersionInfo.FileVersion } catch {}
}

$Install = $false
if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Desktop-Version: $installedVersion | Lokal: $localVersion" -Level 'INFO'
    try { $Install = [version]$installedVersion -lt [version]$localVersion } catch { $Install = $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht auf Desktop installiert." -Level 'INFO'
}

Write-Host ""

# Copy to desktop if needed
if ($Install -or $InstallationFlag) {
    Write-Host "WinReducerEX100 wird kopiert" -ForegroundColor Cyan
    Write-DeployLog -Message "Starte Kopiervorgang: $InstallationFolder -> $destinationPath" -Level 'INFO'
    if (Test-Path $InstallationExe) {
        try {
            Copy-Item -Path $InstallationFolder -Destination $destinationPath -Recurse -Force
            Write-DeployLog -Message "Kopiervorgang erfolgreich." -Level 'SUCCESS'
        } catch {
            Write-DeployLog -Message "Fehler beim Kopieren: $_" -Level 'ERROR'
        }
    } else {
        Write-DeployLog -Message "Quellpfad nicht gefunden: $InstallationExe" -Level 'ERROR'
    }
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
