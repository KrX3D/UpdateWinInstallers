param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Adobe Acrobat"
$ScriptType = "Update"

# DeployToolkit helpers
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (Test-Path $dtPath) {
    Import-Module -Name $dtPath -Force -DisableNameChecking -ErrorAction Stop
} else {
    if (Get-Command -Name Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "DeployToolkit nicht gefunden: $dtPath" -Level "WARNING"
    } else {
        Write-Warning "DeployToolkit nicht gefunden: $dtPath"
    }
}

Initialize-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot
Write-DeployLog -Message "Script gestartet mit InstallationFlag: $InstallationFlag" -Level 'INFO'
Write-DeployLog -Message "ProgramName: $ProgramName, ScriptType: $ScriptType" -Level 'DEBUG'

try {
    $config = Import-DeployConfig -ScriptRoot $PSScriptRoot
    $InstallationFolder = $config.InstallationFolder
    $Serverip = $config.Serverip
    $PSHostPath = $config.PSHostPath
} catch {
    Write-Host ""
    Write-Host "Konfigurationsdatei konnte nicht geladen werden." -ForegroundColor "Red"
    Write-DeployLog -Message "Script beendet wegen fehlender Konfiguration: $_" -Level 'ERROR'
    Finalize-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
    exit
}

# Define the path to the Adobe Acrobat Reader installer
$installerPath = "$InstallationFolder\AcroRdrDC*_de_DE.exe"
Write_LogEntry -Message "Installer path gesetzt: $installerPath" -Level "DEBUG"

$versionPattern = 'AcroRdrDCx64(\d+)_de_DE.exe'
Write_LogEntry -Message "Version pattern gesetzt: $versionPattern" -Level "DEBUG"

$installerFile = Get-InstallerFilePath -PathPattern $installerPath

# Check if the installer file exists
if ($installerFile) {
    Write_LogEntry -Message "Installer gefunden: $($installerFile.Name)" -Level "INFO"
    # Extract the version number from the file name
    $fileVersion = Get-InstallerFileVersion -FilePath $installerFile.FullName -FileNameRegex $versionPattern -Source FileName -Convert { param($v) Convert-AdobeToVersion $v }
    Write_LogEntry -Message "Lokale Installationsdatei Version extrahiert: $fileVersion" -Level "DEBUG"

    # Check if there is a newer version available online
    #https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html
    #https://helpx.adobe.com/de/acrobat/release-note/release-notes-acrobat-reader.html

    $webPageUrl = "https://it-blogger.net/adobe-reader-offline-installer-fuer-windows-und-macos/"
    $latestVersionPattern = 'Adobe Acrobat Reader 64-bit Version (\d+\.\d+\.\d+).*?Windows'

    $onlineInfo = Get-OnlineVersionInfo -Url $webPageUrl -Regex $latestVersionPattern -Context $ProgramName -Transform { param($v) if ($v) { $v -replace '\.' -replace '^..' } else { $null } }

    if ($onlineInfo.Content) {
        Write_LogEntry -Message "Webseite für Versionsprüfung abgerufen: $webPageUrl" -Level "DEBUG"
        Write_LogEntry -Message "Starte Extraktion der Online-Version mit Pattern: $latestVersionPattern" -Level "DEBUG"

        $latestVersion = $onlineInfo.Version
        $latestVersionObj = if ($latestVersion) { Convert-AdobeToVersion $latestVersion } else { $null }
        if ($latestVersionObj) {
            Write_LogEntry -Message "Online Version gefunden: $latestVersionObj (raw=$latestVersion)" -Level "DEBUG"
            Write-Host ""
            Write-Host "Lokale Version: $fileVersion" -ForegroundColor "Cyan"
            Write-Host "Online Version: $latestVersionObj" -ForegroundColor "Cyan"
            Write-Host ""

            if ($latestVersionObj -gt $fileVersion) {
                Write_LogEntry -Message "Neue Version $latestVersionObj verfügbar. Starte Download." -Level "INFO"

                # Construct the download URL for the offline installer
                $downloadUrl = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$latestVersion/AcroRdrDCx64$latestVersion`_de_DE.exe"
                $downloadPath = "$InstallationFolder\AcroRdrDCx64$latestVersion`_de_DE.exe"
                Write_LogEntry -Message "Download URL konstruiert: $downloadUrl" -Level "DEBUG"
                Write_LogEntry -Message "Download Pfad gesetzt: $downloadPath" -Level "DEBUG"

                $downloadOk = Invoke-InstallerDownload -Url $downloadUrl -OutFile $downloadPath -Context $ProgramName
                if ($downloadOk -and (Resolve-DownloadedInstaller -DownloadedFile $downloadPath -RemovePattern $installerPath -ReplaceOld -Context $ProgramName)) {
                    Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor "Green"
                    Write_LogEntry -Message "$ProgramName wurde aktualisiert: $downloadPath" -Level "SUCCESS"
                } else {
                    Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "Red"
                    Write_LogEntry -Message "Download fehlgeschlagen: $downloadPath" -Level "ERROR"
                }
            } else {
                Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
                Write_LogEntry -Message "Kein Online Update verfügbar. Lokale Version $fileVersion ist aktuell." -Level "INFO"
            }
        } else {
            Write_LogEntry -Message "Konnte Online-Version nicht aus der Webseite extrahieren." -Level "WARNING"
        }
    } else {
        Write_LogEntry -Message "Fehler beim Abrufen der Webseite $webPageUrl" -Level "ERROR"
    }
} else {
    #Write-Host "No Adobe Acrobat Reader installer found in the specified path." -ForegroundColor "Red"
    Write_LogEntry -Message "Kein Installer im Pfad gefunden: $installerPath" -Level "WARNING"
}

Write-Host ""

#Check Installed Version / Install if needed
try {
    $latestInstaller = Get-InstallerFilePath -PathPattern $installerPath
    $localVersion = if ($latestInstaller) { Get-InstallerFileVersion -FilePath $latestInstaller.FullName -FileNameRegex $versionPattern -Source FileName -Convert { param($v) Convert-AdobeToVersion $v } } else { $null }
    Write_LogEntry -Message "Lokale Installationsdatei Version (für Vergleiche) ist: $localVersion" -Level "DEBUG"
} catch {
    $localVersion = $null
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Installationsdatei Version: $_" -Level "ERROR"
}

$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.Version } else { $null }
$localVersionObj = if ($localVersion -is [version]) { $localVersion } else { Convert-AdobeToVersion $localVersion }

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor "Green"
    Write-Host "	Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersionObj" -ForegroundColor "Cyan"
    Write_LogEntry -Message "$ProgramName ist installiert. Installierte Version: $installedVersion; Installationsdatei Version: $localVersionObj" -Level "INFO"
} else {
    Write_LogEntry -Message "$ProgramName wurde nicht in der Registrierung gefunden. Setze Install=false" -Level "DEBUG"
}

$state = Compare-VersionState -InstalledVersion $installedVersion -InstallerVersion $localVersionObj -Context $ProgramName
if ($state.UpdateRequired) {
    Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "magenta"
    Write_LogEntry -Message "Veraltete Version erkannt. Update wird gestartet." -Level "INFO"
}
Write-Host ""

#Install if needed

$installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\AdobeDcInstall.ps1"
$started = Invoke-InstallDecision -PSHostPath $PSHostPath -InstallScript $installScript -InstallationFlag:$InstallationFlag -InstallRequired:$state.UpdateRequired -Context $ProgramName
Write-Host ""

Finalize-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
