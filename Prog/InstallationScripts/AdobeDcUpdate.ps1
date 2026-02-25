param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Adobe Acrobat"
$ScriptType = "Update"

# DeployToolkit helpers
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (Test-Path $dtPath) {
    Import-Module -Name $dtPath -Force -ErrorAction Stop
} else {
    if (Get-Command -Name Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "DeployToolkit nicht gefunden: $dtPath" -Level "WARNING"
    } else {
        Write-Warning "DeployToolkit nicht gefunden: $dtPath"
    }
}

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot
Write-DeployLog -Message "Script gestartet mit InstallationFlag: $InstallationFlag" -Level 'INFO'
Write-DeployLog -Message "ProgramName: $ProgramName, ScriptType: $ScriptType" -Level 'DEBUG'

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

# Define the path to the Adobe Acrobat Reader installer
$installerPath = "$InstallationFolder\AcroRdrDC*_de_DE.exe"
Write_LogEntry -Message "Installer path gesetzt: $installerPath" -Level "DEBUG"

$versionPattern = 'AcroRdrDC(?:x64)?(\d+)_(?:de_DE|en_US|MUI)\.exe'
Write_LogEntry -Message "Version pattern gesetzt: $versionPattern" -Level "DEBUG"

$localInstallerInfo = Get-InstallerVersionForComparison -PathPattern $installerPath -FileNameRegex $versionPattern -Convert { param($v) Convert-AdobeToVersion $v } -Context $ProgramName -Description 'Lokale Installationsdatei Version'
$installerFile = $localInstallerInfo.InstallerFile
$fileVersion = $localInstallerInfo.Version

# Check if the installer file exists
if ($installerFile) {
    Write_LogEntry -Message "Installer gefunden: $($installerFile.Name)" -Level "INFO"

    # Check if there is a newer version available online
    #https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html
    #https://helpx.adobe.com/de/acrobat/release-note/release-notes-acrobat-reader.html

    $webPageUrl = "https://it-blogger.net/adobe-reader-offline-installer-fuer-windows-und-macos/"
    $latestVersionPattern = @(
        'Version\s+([0-9]{4}\.[0-9]{3}\.[0-9]{5})',
        'AcroRdrDCx64([0-9]{8,10})_(?:de_DE|en_US|MUI)\.exe',
        'AcroRdrDC([0-9]{8,10})_(?:de_DE|en_US|MUI)\.exe'
    )

    $onlineInfo = Get-OnlineVersionInfo -Url $webPageUrl -Regex $latestVersionPattern -Context $ProgramName -Transform {
        param($v)
        if (-not $v) { return $null }
        if ($v -match '^\d{8,10}$') { return $v }

        $versionObj = Convert-AdobeToVersion $v
        if (-not $versionObj) { return $null }
        return Convert-AdobeVersionToDigits $versionObj
    }

    if ($onlineInfo.Content) {
        Write_LogEntry -Message "Webseite für Versionsprüfung abgerufen: $webPageUrl" -Level "DEBUG"
        Write_LogEntry -Message "Starte Extraktion der Online-Version mit Pattern: $($latestVersionPattern -join ' | ')" -Level "DEBUG"

        $latestVersion = $onlineInfo.Version
        $latestVersionObj = if ($latestVersion) { Convert-AdobeToVersion $latestVersion } else { $null }
        Write_LogEntry -Message "Online Rohversion: $latestVersion; Vergleichsversion: $latestVersionObj" -Level "DEBUG"

        Write-Host ""
        Write-Host "Lokale Version: $fileVersion" -ForegroundColor "Cyan"
        Write-Host "Online Version: $latestVersionObj" -ForegroundColor "Cyan"
        Write-Host ""

        if ($latestVersionObj) {
            Write_LogEntry -Message "Online Version gefunden: $latestVersionObj (raw=$latestVersion)" -Level "DEBUG"
            if ($latestVersionObj -gt $fileVersion) {
                Write_LogEntry -Message "Neue Version $latestVersionObj verfügbar. Starte Download." -Level "INFO"

                # Construct the download URL for the offline installer
                $downloadUrl = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$latestVersion/AcroRdrDCx64$latestVersion`_de_DE.exe"
                $downloadPath = "$InstallationFolder\AcroRdrDCx64$latestVersion`_de_DE.exe"
                Write_LogEntry -Message "Download URL konstruiert: $downloadUrl" -Level "DEBUG"
                Write_LogEntry -Message "Download Pfad gesetzt: $downloadPath" -Level "DEBUG"

                [void](Invoke-InstallerDownload -Url $downloadUrl -OutFile $downloadPath -Context $ProgramName -ConfirmDownload -ReplaceOld -RemoveFiles @($installerFile.FullName) -EmitHostStatus -SuccessHostMessage "$ProgramName wurde aktualisiert.." -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -SuccessLogMessage "$ProgramName wurde aktualisiert: $downloadPath" -FailureLogMessage "Download fehlgeschlagen: $downloadPath")
            } else {
                Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
                Write_LogEntry -Message "Kein Online Update verfügbar. Lokale Version $fileVersion ist aktuell." -Level "INFO"
            }
        } else {
            Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
            Write_LogEntry -Message "Online-Version gefunden, konnte aber nicht in Vergleichsformat umgewandelt werden (raw=$latestVersion)." -Level "WARNING"
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
$comparisonInstaller = Get-InstallerVersionForComparison -PathPattern $installerPath -FileNameRegex $versionPattern -Convert { param($v) Convert-AdobeToVersion $v } -Context $ProgramName -Description 'Lokale Installationsdatei Version (für Vergleiche)'
$localVersion = $comparisonInstaller.Version

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
Show-VersionStateSummary -State $state -ProgramName $ProgramName -Context $ProgramName | Out-Null

#Install if needed

$installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\AdobeDcInstall.ps1"
$started = Invoke-InstallDecision -PSHostPath $PSHostPath -InstallScript $installScript -InstallationFlag:$InstallationFlag -InstallRequired:$state.UpdateRequired -Context $ProgramName
Write-Host ""

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
