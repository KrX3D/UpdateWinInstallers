param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Advanced Port Scanner"
$ScriptType = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

Write_LogEntry -Message "ProgramName gesetzt auf: $($ProgramName)" -Level "DEBUG"

$localFilePath = "$InstallationFolder\Advanced_Port_Scanner_*.exe"
Write_LogEntry -Message "Suche lokale Installationsdatei mit Pattern: $($localFilePath)" -Level "DEBUG"

$webPageUrl = "https://www.advanced-port-scanner.com/de/"
Write_LogEntry -Message "Webseite gesetzt auf: $($webPageUrl)" -Level "DEBUG"

# Get the local file version
$localFile = Get-ChildItem -Path $localFilePath | Select-Object -Last 1
if ($null -ne $localFile) {
    Write_LogEntry -Message "Lokale Datei gefunden: $($localFile.FullName)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine lokale Datei gefunden für Pattern: $($localFilePath)" -Level "WARNING"
}

$localVersion = $null
if ($null -ne $localFile) {
    try {
        $localVersion = (Get-ItemProperty -Path $localFile.FullName).VersionInfo.FileVersion
        Write_LogEntry -Message "Lokale Dateiversion ermittelt: $($localVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Auslesen der lokalen Dateiversion für $($localFile.FullName): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Lokale Version kann nicht ermittelt werden, da keine Datei vorhanden ist." -Level "WARNING"
}

if ($null -ne $localVersion) {
    $versionParts = $localVersion.Split('.')
    $trimmedVersion = ($versionParts[0], $versionParts[1], $versionParts[2] -replace '^0+(\d)', '$1') -join '.'
    $localVersion = $trimmedVersion
    Write_LogEntry -Message "Bereinigte lokale Version: $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Lokale Version ist leer/null." -Level "DEBUG"
}

# Retrieve the web page content
Write_LogEntry -Message "Rufe Webseite $($webPageUrl) ab." -Level "INFO"
$webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing
if ($null -ne $webPageContent) {
    Write_LogEntry -Message "Webseiteninhalt empfangen. Länge Content: $($webPageContent.Content.Length)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine Daten von $($webPageUrl) erhalten." -Level "WARNING"
}

# Extract the download link from the web page content
$downloadLinkPattern = '<a href="(https://download\.advanced-port-scanner\.com/download/files/Advanced_Port_Scanner_[\d\.]+\.exe)"'
$downloadLinkMatch = [regex]::Match($webPageContent.Content, $downloadLinkPattern)

if ($downloadLinkMatch.Success) {
    $downloadLink = $downloadLinkMatch.Groups[1].Value
    Write_LogEntry -Message "Downloadlink gefunden: $($downloadLink)" -Level "INFO"

    # Extract the version number from the download link
    $versionPattern = 'Advanced_Port_Scanner_(\d+\.\d+\.\d+)'
    $versionMatch = [regex]::Match($downloadLink, $versionPattern)

    if ($versionMatch.Success) {
        $onlineVersion = $versionMatch.Groups[1].Value
        Write_LogEntry -Message "Online-Version extrahiert: $($onlineVersion)" -Level "DEBUG"

		Write-Host ""
		Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
		Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
		Write-Host ""
		
        # Compare the local and online file versions
        if ($onlineVersion -gt $localVersion) {
            Write_LogEntry -Message "Online-Version $($onlineVersion) ist größer als lokale Version $($localVersion). Update wird vorbereitet." -Level "INFO"

            # Extract the file name from the download link
            $fileName = [System.IO.Path]::GetFileName($downloadLink)
            Write_LogEntry -Message "Dateiname aus Downloadlink: $($fileName)" -Level "DEBUG"

            # Download the updated installer
            $downloadedFilePath = "$InstallationFolder\$fileName"
            Write_LogEntry -Message "Starte Download von $($downloadLink) nach $($downloadedFilePath)" -Level "INFO"
			
            #Invoke-WebRequest -Uri $downloadLink -OutFile $downloadedFilePath
			$webClient = New-Object System.Net.WebClient
			[void](Invoke-DownloadFile -Url $downloadLink -OutFile $downloadedFilePath)
			$webClient.Dispose()
            Write_LogEntry -Message "Download abgeschlossen (temporär, Überprüfung folgt): $($downloadedFilePath)" -Level "DEBUG"
			
			# Check if the file was completely downloaded
			if (Test-Path $downloadedFilePath) {
                Write_LogEntry -Message "Download bestätigt, Datei vorhanden: $($downloadedFilePath)" -Level "SUCCESS"
				# Remove the old installer
                if ($null -ne $localFile) {
                    try {
                        Remove-Item -Path $localFile.FullName -Force
                        Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFile.FullName)" -Level "INFO"
                    } catch {
                        Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($localFile.FullName): $($_)" -Level "ERROR"
                    }
                } else {
                    Write_LogEntry -Message "Keine alte Installationsdatei zum Entfernen gefunden." -Level "DEBUG"
                }

				Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                Write_LogEntry -Message "$($ProgramName) Update abgeschlossen. Neue Datei: $($downloadedFilePath)" -Level "SUCCESS"
			} else {
				Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
                Write_LogEntry -Message "Download fehlgeschlagen, Datei nicht vorhanden: $($downloadedFilePath)" -Level "ERROR"
			}
        } else {
			Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
            Write_LogEntry -Message "Kein Update erforderlich. Online-Version $($onlineVersion) <= lokale Version $($localVersion)" -Level "INFO"
        }
    } else {
        Write_LogEntry -Message "Versionsinformation nicht im Downloadlink gefunden: $($downloadLink)" -Level "WARNING"
        #Write-Host "Version information not found in the download link."
    }
} else {
    Write_LogEntry -Message "Downloadlink auf der Webseite nicht gefunden: $($webPageUrl)" -Level "WARNING"
    #Write-Host "Download link not found on the website."
}

Write-Host ""
Write_LogEntry -Message "Erster Abschnitt (Prüfung/Download) beendet." -Level "DEBUG"

#Check Installed Version / Install if neded
Write_LogEntry -Message "Ermittle erneut lokale Installationsdatei für Check: $($localFilePath)" -Level "DEBUG"
$localFile = Get-ChildItem -Path $localFilePath | Select-Object -Last 1
if ($null -ne $localFile) {
    Write_LogEntry -Message "Lokale Datei für Check gefunden: $($localFile.FullName)" -Level "DEBUG"
    try {
        $localVersion = (Get-ItemProperty -Path $localFile.FullName).VersionInfo.FileVersion
        Write_LogEntry -Message "Lokale Dateiversion (erneut) ermittelt: $($localVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Auslesen der lokalen Dateiversion (erneut) für $($localFile.FullName): $($_)" -Level "ERROR"
    }
    $versionParts = $localVersion.Split('.')
    $trimmedVersion = ($versionParts[0], $versionParts[1], $versionParts[2] -replace '^0+(\d)', '$1') -join '.'
    $localVersion = $trimmedVersion
    Write_LogEntry -Message "Bereinigte lokale Version (erneut): $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine lokale Datei für Installations-Check gefunden." -Level "WARNING"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Prüfe Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registry-Pfad vorhanden: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht vorhanden: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	Write_LogEntry -Message "Programm installiert. Installierte Version: $($installedVersion). Lokale Installationsdatei Version: $($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Installationsentscheidung: Update erforderlich (Install = $($Install))." -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Keine Aktion erforderlich (Install = $($Install))." -Level "INFO"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install gesetzt auf $($Install) (installierte Version > lokale Version)." -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
    Write_LogEntry -Message "Programm $($ProgramName) nicht in Registry gefunden. Install = $($Install)" -Level "INFO"
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    $targetFile = "$Serverip\Daten\Prog\InstallationScripts\Installation\AdvancedPortScannerInstallation.ps1"
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte externes Installationsskript: $($targetFile) mittels $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\AdvancedPortScannerInstallation.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Aufruf externes Skript beendet: $($targetFile)" -Level "DEBUG"
}
elseif($Install -eq $true){
    $targetFile2 = "$Serverip\Daten\Prog\InstallationScripts\Installation\GitInsAdvancedPortScannerInstallationtall.ps1"
    Write_LogEntry -Message "Install=true. Starte externes Installationsskript: $($targetFile2) mittels $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\GitInsAdvancedPortScannerInstallationtall.ps1"
    Write_LogEntry -Message "Aufruf externes Skript beendet: $($targetFile2)" -Level "DEBUG"
}
Write-Host ""

# ===== Logger-Footer (BEGIN) =====
Write_LogEntry -Message "Script beendet." -Level "INFO"
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
# ===== Logger-Footer (END) =====
