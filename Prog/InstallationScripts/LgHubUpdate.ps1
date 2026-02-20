param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Logitech G HUB"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PSScriptRoot: $($PSScriptRoot)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

#"C:\Program Files\Google\Chrome\Application\chrome.exe" --headless --run-all-compositor-stages-before-draw --print-to-pdf=C:\Users\KrX\Downloads\output.pdf "https://support.logi.com/hc/en-us/articles/360025298133" --virtual-time-budget=60000
#"C:\Program Files\Google\Chrome\Application\chrome.exe" --headless --run-all-compositor-stages-before-draw --virtual-time-budget=60000  --dump-dom "https://support.logi.com/hc/en-us/articles/360025298133" > source.html

$InstallationFolder = "$NetworkShareDaten\Treiber\LgHub"
$filenamePattern = "lghub_installer*.exe"

Write_LogEntry -Message "InstallationFolder: $($InstallationFolder); FilenamePattern: $($filenamePattern)" -Level "DEBUG"

# Get the latest local file matching the pattern
$latestFile = Get-InstallerFilePath -Directory $InstallationFolder -Filter $filenamePattern
Write_LogEntry -Message "Gefundene lokale Datei (latestFile exists): $($([bool]$latestFile)); Pfad: $($($latestFile.FullName -as [string]))" -Level "DEBUG"

if ($latestFile) {
    # Get the version information from the local installer
    $localVersion = $latestFile.Name -replace 'lghub_installer_|\.exe', ''
    Write_LogEntry -Message "Lokale Version aus Dateiname extrahiert: $($localVersion); Datei: $($latestFile.FullName)" -Level "INFO"

	#Chrome
	$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
	#$headlessArgs = "--headless --run-all-compositor-stages-before-draw --virtual-time-budget=60000"
	#$headlessArgs = "--headless=old --run-all-compositor-stages-before-draw --virtual-time-budget=60000"
	$headlessArgsOld = "--headless=old --run-all-compositor-stages-before-draw --virtual-time-budget=60000"
	$headlessArgsNew = "--headless --run-all-compositor-stages-before-draw --virtual-time-budget=60000"

	$tempFolder = [System.IO.Path]::GetTempPath()
	$outputPath = Join-Path $tempFolder "LGHUB.html"
	Write_LogEntry -Message "TempFolder: $($tempFolder); OutputPath: $($outputPath)" -Level "DEBUG"

    # Check online for a newer version
	$latestVersionUrl = "https://support.logi.com/hc/en-us/articles/360025298133"
    $latestVersion = "0"  # Default fallback value
    $downloadUrl = "https://download01.logi.com/web/ftp/pub/techsupport/gaming/lghub_installer.exe"  # Default fallback URL
	Write_LogEntry -Message "Online-Check URL: $($latestVersionUrl); Fallback DownloadURL: $($downloadUrl)" -Level "DEBUG"

	# Check if Chrome executable exists
	if (Test-Path $chromePath) {
        Write_LogEntry -Message "Chrome gefunden: $($chromePath)" -Level "INFO"
		#Start-Process -FilePath $chromePath -ArgumentList "$headlessArgs --dump-dom  $latestVersionUrl" -RedirectStandardOutput $outputPath -NoNewWindow -Wait

		# Function to execute the Chrome command
		function Run-ChromeDump {
			param (
				[string]$Arguments
			)
            try {
                Start-Process -FilePath $chromePath -ArgumentList "$Arguments --dump-dom $latestVersionUrl" -RedirectStandardOutput $outputPath -NoNewWindow -Wait -ErrorAction Stop
                Write_LogEntry -Message "Chrome Aufruf erfolgreich mit Argumenten: $($Arguments)" -Level "DEBUG"
            }
            catch {
                Write_LogEntry -Message "Chrome execution failed: $($($_.Exception.Message))" -Level "ERROR"
                Write-Host "Chrome execution failed: $($_.Exception.Message)" -ForegroundColor "Red"
            }
        }

		# Run with --headless=old first
		Write_LogEntry -Message "Starte Chrome Dump mit --headless=old" -Level "DEBUG"
		Run-ChromeDump -Arguments $headlessArgsOld

		# Check if the file is 0 KB
        if ((Test-Path $outputPath) -and (Get-Item $outputPath).Length -eq 0) {
			Write_LogEntry -Message "Output-Datei $($outputPath) 0 KB; erneuter Versuch mit --headless" -Level "WARNING"
			#Write-Host "The file is 0 KB. Retrying with --headless..."
			Run-ChromeDump -Arguments $headlessArgsNew
		}

        # Check if outputPath exists and has content
        if ((Test-Path $outputPath) -and (Get-Item $outputPath).Length -gt 0) {
            try {
				$htmlContent = Get-Content $outputPath -Raw
                Write_LogEntry -Message "HTML-Inhalt erfolgreich geladen aus $($outputPath)" -Level "DEBUG"
                $versionPattern = "<b><span>Software Version: </span></b>(\d+\.\d+\.\d+)"
                $versionMatch = [regex]::Match($htmlContent, $versionPattern)

                if ($versionMatch.Success) {
                    $latestVersion = $versionMatch.Groups[1].Value
                    Write_LogEntry -Message "Online-Version erfolgreich abgerufen: $($latestVersion)" -Level "INFO"
                    #Write-Host "Online-Version erfolgreich abgerufen: $latestVersion" -ForegroundColor "Green"
                } else {
                    Write_LogEntry -Message "Versionsnummer nicht im HTML gefunden; Verwende Fallback-Version: $($latestVersion)" -Level "WARNING"
                    #Write-Host "Versionsnummer nicht im HTML gefunden. Verwende Fallback-Version." -ForegroundColor "Yellow"
                }

                $linkPattern = '<a class="download-button" href="(.*?)" target="_blank">Download Now</a>'
                
                $linkMatch = [regex]::Match($htmlContent, $linkPattern)

                if ($linkMatch.Success) {
                    $downloadUrl = $linkMatch.Groups[1].Value
                    Write_LogEntry -Message "Download-URL aus HTML extrahiert: $($downloadUrl)" -Level "INFO"
                } else {
                    Write_LogEntry -Message "Downloadlink im HTML nicht gefunden; Verwende Fallback-DownloadURL: $($downloadUrl)" -Level "WARNING"
                }
            }
            catch {
                Write_LogEntry -Message "Fehler beim Verarbeiten der HTML-Datei: $($($_.Exception.Message))" -Level "ERROR"
                Write-Host "Fehler beim Verarbeiten der HTML-Datei: $($_.Exception.Message)" -ForegroundColor "Red"
            }
        } else {
            Write_LogEntry -Message "Kein oder leeres Output-File $($outputPath) nach Chrome-Aufruf; Fallback-Werte werden verwendet." -Level "WARNING"
            Write-Host "Chrome konnte keine gültigen Daten abrufen. Verwende Fallback-Werte." -ForegroundColor "Yellow"
        }
    } else {
        Write_LogEntry -Message "Chrome nicht gefunden unter Pfad: $($chromePath); Fallback-Werte werden verwendet." -Level "WARNING"
        Write-Host "Chrome nicht gefunden. Verwende Fallback-Werte." -ForegroundColor "Yellow"
    }

	Write_LogEntry -Message "Lokale Version: $($localVersion); Online Version (erkannt): $($latestVersion)" -Level "INFO"

	Write-Host ""
	Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
	Write-Host ""

    # Only proceed with version comparison if we have valid version strings
    if ($latestVersion -ne "0" -and $localVersion -match '^\d+\.\d+\.\d+$' -and $latestVersion -match '^\d+\.\d+\.\d+$') {
        try {
			#Versionsnummern aufspalten
			$localComponents = $localVersion -split '\.'
			$onlineComponents = $latestVersion -split '\.'

			# Convert the components to integers for numerical comparison
			$localYear = [int]$localComponents[0]
			$localMonth = [int]$localComponents[1]
			$localBuild = [int]$localComponents[2]

			$onlineYear = [int]$onlineComponents[0]
			$onlineMonth = [int]$onlineComponents[1]
			$onlineBuild = [int]$onlineComponents[2]

            Write_LogEntry -Message "Vergleiche Versionen Lokal: $($localYear).$($localMonth).$($localBuild) vs Online: $($onlineYear).$($onlineMonth).$($onlineBuild)" -Level "DEBUG"

			if (
				$onlineYear -gt $localYear -or
				(($onlineYear -eq $localYear) -and ($onlineMonth -gt $localMonth)) -or
				(($onlineYear -eq $localYear) -and ($onlineMonth -eq $localMonth) -and ($onlineBuild -gt $localBuild))
			) {
		        $downloadPath = Join-Path -Path $InstallationFolder -ChildPath "lghub_installer_$latestVersion.exe"
				Write_LogEntry -Message "Update verfügbar. Ziel-Downloadpfad: $($downloadPath)" -Level "INFO"
				
                try {
                    Write-Host "Lade neue Version herunter..." -ForegroundColor "Yellow"
                    Write_LogEntry -Message "Starte Download von $($downloadUrl) nach $($downloadPath)" -Level "INFO"
		        	#Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath
                    $webClient = New-Object System.Net.WebClient
                    [void](Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath)
                    $webClient.Dispose()
                    Write_LogEntry -Message "Download abgeschlossen; prüfe Existenz: $($downloadPath)" -Level "DEBUG"
                    
                    # Check if the file was completely downloaded
                    if (Test-Path $downloadPath) {
                        # Remove the old installer
                        Remove-Item -Path $latestFile.FullName -Force
                        Write_LogEntry -Message "Alte Installationsdatei entfernt: $($latestFile.FullName)" -Level "DEBUG"

                        Write-Host "$ProgramName wurde aktualisiert." -foregroundcolor "Green"
                        Write_LogEntry -Message "$($ProgramName) wurde aktualisiert auf $($latestVersion)" -Level "SUCCESS"
                    } else {
                        Write_LogEntry -Message "Download ist fehlgeschlagen; Datei nicht gefunden: $($downloadPath)" -Level "ERROR"
                        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "Red"
                    }
                }
                catch {
                    Write_LogEntry -Message "Download-Fehler: $($($_.Exception.Message))" -Level "ERROR"
                    Write-Host "Download-Fehler: $($_.Exception.Message)" -foregroundcolor "Red"
                }
            } else {
                Write_LogEntry -Message "Kein Online Update verfügbar. Lokal ist aktuell: Online $($latestVersion) <= Lokal $($localVersion)" -Level "INFO"
                Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -foregroundcolor "DarkGray"
            }
        }
        catch {
            Write_LogEntry -Message "Fehler beim Vergleichen der Versionen: $($($_.Exception.Message))" -Level "ERROR"
            Write-Host "Fehler beim Vergleichen der Versionen: $($_.Exception.Message)" -foregroundcolor "Red"
            Write-Host "Überspringe Versionsprüfung." -foregroundcolor "Yellow"
        }
    } else {
        Write_LogEntry -Message "Ungültige Versionsdaten - überspringe Online-Update-Prüfung. Lokal: $($localVersion); Online: $($latestVersion)" -Level "WARNING"
        Write-Host "Ungültige Versionsdaten - überspringe Online-Update-Prüfung." -foregroundcolor "Yellow"
    }
	
    # Remove the HTML file if it exists
    if (Test-Path $outputPath) {
        Remove-Item $outputPath -Force
        Write_LogEntry -Message "Temporäre HTML-Datei entfernt: $($outputPath)" -Level "DEBUG"
    }
} else {
    Write_LogEntry -Message "Keine lokale lghub_installer Datei im Ordner $($InstallationFolder) gefunden; überspringe Update-Check." -Level "WARNING"
}

Write_LogEntry -Message "Ermittle aktuellste lokale Datei nochmals für Installationsprüfung." -Level "DEBUG"

Write-Host ""

#Check Installed Version / Install if needed
$FoundFile = Get-InstallerFilePath -Directory $InstallationFolder -Filter $filenamePattern

if ($FoundFile) {
	$localVersion = $FoundFile.Name -replace 'lghub_installer_|\.exe', ''
    Write_LogEntry -Message "Für Installationsprüfung gefundene Datei: $($FoundFile.FullName); Lokale Version: $($localVersion)" -Level "DEBUG"

	#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

    $RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    Write_LogEntry -Message "Zu prüfende Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

    $Path = foreach ($RegPath in $RegistryPaths) {
        if (Test-Path $RegPath) {
            Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
        }
    }
    Write_LogEntry -Message "Registry-Abfrage Ergebnis vorhanden: $($([bool]$Path))" -Level "DEBUG"

    if ($null -ne $Path) {
        $installedVersion = $Path.DisplayVersion | Select-Object -First 1
        Write_LogEntry -Message "Gefundene installierte Version: $($installedVersion); Installationsdatei Version: $($localVersion)" -Level "INFO"
        Write-Host "$ProgramName ist installiert." -foregroundcolor "Green"
        Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
        Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"

        try {
            if ([version]$installedVersion -lt [version]$localVersion) {
                Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist älter als lokale Datei ($($localVersion)): Install = True" -Level "INFO"
                Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "Magenta"
                $Install = $true
            } elseif ([version]$installedVersion -eq [version]$localVersion) {
                Write_LogEntry -Message "Installierte Version entspricht lokaler Datei: $($installedVersion): Install = False" -Level "DEBUG"
                Write-Host "		Installierte Version ist aktuell." -ForegroundColor "DarkGray"
                $Install = $false
            } else {
                Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Datei ($($localVersion)): Install = False" -Level "WARNING"
                $Install = $false
            }
        }
        catch {
            Write_LogEntry -Message "Fehler beim Vergleichen der installierten Version: $($($_.Exception.Message))" -Level "ERROR"
            Write-Host "		Fehler beim Vergleichen der installierten Version: $($_.Exception.Message)" -foregroundcolor "Red"
            $Install = $false
        }
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
        Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden; Install = $($Install)" -Level "INFO"
    }
}

Write_LogEntry -Message "Installationsentscheidung: Install = $($Install)" -Level "DEBUG"
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript mit -InstallationFlag: $($Serverip)\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte externes Installations-Skript (Update): $($Serverip)\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1"
    Write_LogEntry -Message "Externes Installations-Skript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\LgHubInstallation.ps1" -Level "DEBUG"
}
Write_LogEntry -Message "Script-Ende erreicht. Vor Footer." -Level "INFO"

Write-Host ""

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
