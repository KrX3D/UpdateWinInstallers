param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "NTLite"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

$InstallationFolder = "$InstallationFolder\NTLite"
Write_LogEntry -Message "InstallationFolder gesetzt: $($InstallationFolder)" -Level "DEBUG"

$localFilePath = "$InstallationFolder\NTLite_setup_x64.exe"
$webPageUrl = "https://www.ntlite.com/download/"
$downloadLinkPattern = 'a href="(https:\/\/downloads\.ntlite\.com\/files\/NTLite_setup_x64\.exe)"'
$versionPattern = 'v(\d+\.\d+\.\d+)'
Write_LogEntry -Message "Lokaler Datei-Pfad: $($localFilePath); WebPageUrl: $($webPageUrl)" -Level "DEBUG"
Write_LogEntry -Message "DownloadLinkPattern: $($downloadLinkPattern); VersionPattern: $($versionPattern)" -Level "DEBUG"

# Get the local file version
try {
    $fileVersionInfo = Get-Item $localFilePath | Get-ItemProperty -Name VersionInfo
    $localVersion = $fileVersionInfo.VersionInfo.FileVersion
    $localVersion = $localVersion -split '\.' | Select-Object -First 3
    $localVersion = $localVersion -join '.'
    Write_LogEntry -Message "Lokale Dateiversion ermittelt: $($localVersion) für Datei $($localFilePath)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Dateiversion für $($localFilePath): $($($_.Exception.Message))" -Level "WARNING"
    $localVersion = $null
}

# Retrieve the web page content
try {
    $webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing
    Write_LogEntry -Message "Webseite abgerufen: $($webPageUrl); InhaltLaenge: $($webPageContent.Content.Length)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen der Webseite $($webPageUrl): $($($_.Exception.Message))" -Level "ERROR"
    $webPageContent = $null
}

if ($webPageContent) {
    # Extract the download link for NTLite_setup_x64.exe
    $match = [regex]::Match($webPageContent.Content, $downloadLinkPattern)
    if ($match.Success) {
        $downloadLink = $match.Groups[1].Value
        Write_LogEntry -Message "Downloadlink extrahiert: $($downloadLink)" -Level "INFO"

        # Extract the name of the file from the download link
        $downloadFileName = Split-Path -Leaf $downloadLink
        Write_LogEntry -Message "Download-Dateiname bestimmt: $($downloadFileName)" -Level "DEBUG"

        # Extract the online version number
        $versionMatch = [regex]::Match($webPageContent.Content, $versionPattern)
        if ($versionMatch.Success) {
            $onlineVersion = $versionMatch.Groups[1].Value
            try {
                $remoteFileVersion = [version]$onlineVersion
            } catch {
                $remoteFileVersion = $null
                Write_LogEntry -Message "Fehler beim Parsen der Online-Versionsnummer: $($onlineVersion)" -Level "WARNING"
            }
            Write_LogEntry -Message "Online-Version extrahiert: $($onlineVersion)" -Level "DEBUG"

            Write-Host ""
            Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
            Write-Host "Online Version: $remoteFileVersion" -foregroundcolor "Cyan"
            Write-Host ""

            # Compare the local and remote file versions
            try {
                if ($remoteFileVersion -gt $localVersion) {
                    Write_LogEntry -Message "Update verfügbar: Online ($($remoteFileVersion)) > Lokal ($($localVersion))" -Level "INFO"

                    $directoryPath = [System.IO.Path]::GetTempPath()
                    # Modify the $downloadPath variable to include the extracted file name
                    $downloadPath = Join-Path $directoryPath $downloadFileName
                    Write_LogEntry -Message "Download-Pfad gesetzt: $($downloadPath)" -Level "DEBUG"

                    # Download the updated installer
                    $webClient = New-Object System.Net.WebClient
                    try {
                        Write_LogEntry -Message "Starte Download von $($downloadLink) nach $($downloadPath)" -Level "INFO"
                        [void](Invoke-DownloadFile -Url $downloadLink -OutFile $downloadPath)
                        Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "DEBUG"
                    } catch {
                        Write_LogEntry -Message "Fehler beim Herunterladen von $($downloadLink): $($($_.Exception.Message))" -Level "ERROR"
                    } finally {
                        $webClient.Dispose()
                    }

                    # Check if the file was completely downloaded
                    if (Test-Path $downloadPath) {
                        try {
                            # Remove the old installer
                            Remove-Item -Path $localFilePath -Force
                            Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFilePath)" -Level "DEBUG"
                        } catch {
                            Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei $($localFilePath): $($($_.Exception.Message))" -Level "WARNING"
                        }

                        try {
                            # Move the downloaded file to the installation folder
                            Move-Item -Path $downloadPath -Destination $installationFolder
                            Write_LogEntry -Message "Heruntergeladene Datei verschoben nach: $($installationFolder)" -Level "SUCCESS"
                        } catch {
                            Write_LogEntry -Message "Fehler beim Verschieben der Datei $($downloadPath) nach $($installationFolder): $($($_.Exception.Message))" -Level "ERROR"
                        }

                        Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                    } else {
                        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
                        Write_LogEntry -Message "Download fehlgeschlagen; Datei nicht gefunden: $($downloadPath)" -Level "ERROR"
                    }
                } else {
                    Write_LogEntry -Message "Kein Online Update verfügbar. Online: $($remoteFileVersion); Lokal: $($localVersion)" -Level "INFO"
                    Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Vergleich der Versionen: $($($_.Exception.Message))" -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Online-Versionsnummer konnte nicht extrahiert werden von $($webPageUrl)" -Level "WARNING"
            #Write-Host "Online version number not found on the website."
        }
    } else {
        Write_LogEntry -Message "Downloadlink nicht auf der Webseite gefunden: $($webPageUrl)" -Level "WARNING"
        #Write-Host "Download link not found on the website."
    }
} else {
    Write_LogEntry -Message "Webseiteninhalt leer oder Fehler beim Abrufen: $($webPageUrl)" -Level "ERROR"
}

Write-Host ""

#Check Installed Version / Install if neded
try {
    $FoundFile = Get-ChildItem $localFilePath
    Write_LogEntry -Message "Gefundene lokale Datei für Prüfung: $($FoundFile.FullName)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Keine lokale Datei für Prüfung gefunden mit Pfad: $($localFilePath)" -Level "WARNING"
    $FoundFile = $null
}
if ($FoundFile) {
    try {
        $localVersion = (Get-Item $FoundFile).VersionInfo.ProductVersion
        Write_LogEntry -Message "Lokale Produktversion ermittelt: $($localVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der Produktversion der Datei $($FoundFile.FullName): $($($_.Exception.Message))" -Level "WARNING"
        $localVersion = $null
    }
} else {
    $localVersion = $null
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Abfrage: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}
Write_LogEntry -Message "Registry-Abfrage durchgeführt; Ergebnis vorhanden: $($([bool]$Path))" -Level "DEBUG"

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write_LogEntry -Message "Gefundene installierte Version in Registry: $($installedVersion)" -Level "INFO"
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (installiere neues Paket)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (Version aktuell)" -Level "DEBUG"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (installierte Version neuer)" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
    Write_LogEntry -Message "$($ProgramName) ist nicht in der Registry gefunden worden (nicht installiert)." -Level "INFO"
}
Write-Host ""
Write_LogEntry -Message "Installationsprüfung abgeschlossen. Install variable: $($Install)" -Level "DEBUG"

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt; rufe Installationsskript mit Flag auf: $($Serverip)\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Installationsskript mit Flag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1" -Level "DEBUG"
}
elseif($Install -eq $true){
    Write_LogEntry -Message "Starte externes Installationsskript (Update): $($Serverip)\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1"
    Write_LogEntry -Message "Externes Installationsskript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\NtLiteInstall.ps1" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
