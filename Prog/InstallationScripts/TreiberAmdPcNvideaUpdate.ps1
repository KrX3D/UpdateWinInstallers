param(
    [switch]$InstallationFlag = $false
)

#https://github.com/Aetopia/NVIDIA-Driver-Package-Downloader
$ProgramName = "NVIDIA Grafiktreiber"
$ScriptType  = "Update"

# === Logger-Header: automatisch eingefügt ===
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Logger\Logger.psm1"

if (Test-Path $modulePath) {
    Import-Module -Name $modulePath -Force -ErrorAction Stop

    if (-not (Get-Variable -Name logRoot -Scope Script -ErrorAction SilentlyContinue)) {
        $logRoot = Join-Path -Path $PSScriptRoot -ChildPath "Log"
    }
    Set_LoggerConfig -LogRootPath $logRoot | Out-Null

    if (Get-Command -Name Initialize_LogSession -ErrorAction SilentlyContinue) {
        Initialize_LogSession -ProgramName $ProgramName -ScriptType $ScriptType | Out-Null #-WriteSystemInfo
    }
}
# === Ende Logger-Header ===

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName initialisiert: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

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

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Konfigurationspfad gesetzt: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei gefunden und importiert: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

$InstallationFolder = "$NetworkShareDaten\Treiber\AMD_PC"
Write_LogEntry -Message "Installations-Ordner: $($InstallationFolder)" -Level "DEBUG"

#Check Pc, Skipp if not AMD PC
$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
Write_LogEntry -Message "Script zur PC-Ermittlung: $($scriptPath)" -Level "DEBUG"

try {
    Write_LogEntry -Message "Aufruf des PC-Ermittlungs-Scripts: $($scriptPath)" -Level "INFO"
	#$PCName = & $scriptPath
	#$PCName = & "powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
	$PCName = & $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File $scriptPath `
		-Verbose:$false

    Write_LogEntry -Message "PCName ermittelt: $($PCName)" -Level "SUCCESS"
		
} catch {
    Write_LogEntry -Message "Fehler beim Laden des Scripts $($scriptPath): $($_)" -Level "ERROR"
    Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
    Exit
}

# Print the pc name
Write-Host "PC Name: $PCName"
Write_LogEntry -Message "PC Name ausgegeben: $($PCName)" -Level "DEBUG"

if($PCName -eq "KrX-AMD-PC"){
	Write_LogEntry -Message "Dieses System entspricht dem Ziel-PC: $($PCName). Starte Treiber-Workflow für $($ProgramName)." -Level "INFO"

	# Define the folder pattern
	$folderPattern = "desktop-win10-win11-64bit-international-nsd-dch-whql$"  # Match folders ending with this pattern
	$versionPattern = '^([\d.]+)-'  # Capture the version number at the beginning of the folder name
	Write_LogEntry -Message "Folder-Pattern: $($folderPattern); Version-Pattern: $($versionPattern)" -Level "DEBUG"
	
	function GetVersionAndDownloadLink {
		param(
			[string]$CurrentLocalVersion
		)
		try {
			Write_LogEntry -Message "Abfrage NVIDIA API für Studio Driver..." -Level "INFO"
			
			$apiUrl = "https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php"
			
			# Parameters for RTX 3090 Ti, Windows 10/11 64-bit, Studio Driver
			$params = @{
				func = "DriverManualLookup"
				psid = "120"           # GeForce RTX 30 Series
				pfid = "929"           # RTX 3090 Ti
				osID = "57"            # Windows 10/11 64-bit
				languageCode = "1033" # English
				beta = "0"
				isWHQL = "1"
				dltype = "-1"
				dch = "1"
				upCRD = "0"
				qnf = "0"
				sort1 = "0"
				numberOfResults = "10"  # Get more results to find available downloads
			}
			
			$queryString = ($params.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
			$fullUrl = "$apiUrl`?$queryString"
			
			$response = Invoke-RestMethod -Uri $fullUrl -Method Get -UseBasicParsing
			
			if ($response.IDS -and $response.IDS.Count -gt 0) {
				# Filter to only show versions newer than local
				$newerVersions = @()
				foreach ($driver in $response.IDS) {
					$ver = $driver.downloadInfo.Version
					if ($CurrentLocalVersion -and [version]$ver -gt [version]$CurrentLocalVersion) {
						$newerVersions += $driver
					}
				}
				
				if ($newerVersions.Count -eq 0) {
					Write_LogEntry -Message "Keine neueren Treiber-Versionen als $CurrentLocalVersion verfügbar" -Level "INFO"
					Write-Host "Keine neueren Treiber-Versionen als $CurrentLocalVersion verfügbar" -ForegroundColor "Green"
					return $null
				}
				
				Write_LogEntry -Message "NVIDIA API hat $($newerVersions.Count) neuere Treiber-Version(en) zurückgegeben (aktuell lokal: $CurrentLocalVersion):" -Level "INFO"
				Write-Host ""
				Write-Host "Verfügbare Updates (aktuell lokal: $CurrentLocalVersion):" -ForegroundColor "Cyan"
				
				# Log only newer versions
				for ($i = 0; $i -lt $newerVersions.Count; $i++) {
					$driverInfo = $newerVersions[$i]
					$ver = $driverInfo.downloadInfo.Version
					$releaseDate = $driverInfo.downloadInfo.ReleaseDateTime
					$constructedUrl = "https://de.download.nvidia.com/Windows/$ver/$ver-desktop-win10-win11-64bit-international-nsd-dch-whql.exe"
					
					Write_LogEntry -Message "  [$($i+1)] Version: $ver | Veröffentlicht: $releaseDate | URL: $constructedUrl" -Level "INFO"
					Write-Host "  [$($i+1)] Version: $ver | Veröffentlicht: $releaseDate" -ForegroundColor "Cyan"
				}
				Write-Host ""
				
				# Try each newer version until we find one with an available download
				foreach ($driver in $newerVersions) {
					$version = $driver.downloadInfo.Version
					
					# Studio Driver uses -nsd suffix (NVIDIA Studio Driver)
					$downloadUrl = "https://de.download.nvidia.com/Windows/$version/$version-desktop-win10-win11-64bit-international-nsd-dch-whql.exe"
					
					Write_LogEntry -Message "Prüfe Download-URL für Version $version..." -Level "DEBUG"
					Write-Host "  -> Prüfe Verfügbarkeit von Version $version..." -ForegroundColor "Yellow"
					
					# Validate that the download URL actually exists
					try {
						$headRequest = Invoke-WebRequest -Uri $downloadUrl -Method Head -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
						
						if ($headRequest.StatusCode -eq 200) {
							Write_LogEntry -Message "NVIDIA API erfolgreich: Version $version gefunden und Download-Link verifiziert" -Level "SUCCESS"
							Write-Host "  [OK] Version $version ist zum Download verfügbar!" -ForegroundColor "Green"
							Write-Host ""
							
							return @{
								VersionNumber = $version
								DownloadLink  = $downloadUrl
							}
						}
					}
					catch {
						Write_LogEntry -Message "Download-Link für Version $version nicht verfügbar (HTTP-Fehler: $($_.Exception.Message)). Versuche nächste Version..." -Level "WARNING"
						Write-Host "  [X] Version $version noch nicht verfügbar" -ForegroundColor "Red"
						continue
					}
				}
				
				Write_LogEntry -Message "Keine verfügbaren Download-Links für neuere Versionen gefunden" -Level "WARNING"
				Write-Host ""
				Write-Host "Keine neueren Treiber zum Download verfügbar." -ForegroundColor "Yellow"
			}
			else {
				Write_LogEntry -Message "Keine Treiber von NVIDIA API zurückgegeben" -Level "ERROR"
			}
			
			return $null
		}
		catch {
			Write_LogEntry -Message "NVIDIA API Fehler: $($_)" -Level "ERROR"
			return $null
		}
	}

	function ExtractExeFile {
		$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

		# Check if 7-Zip is installed by verifying the existence of the executable
		if (Test-Path $sevenZipPath) {
			#Write-Host "7-Zip ist installiert."

			# Define the download path and file name
			$downloadPath = Join-Path $InstallationFolder $downloadFileName

			# Extract the filename without extension to create the target folder
			$fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($downloadFileName)
			$targetFolder = Join-Path $InstallationFolder $fileNameWithoutExtension

			# Create the target folder if it doesn't exist
			if (-not (Test-Path $targetFolder)) {
				New-Item -Path $targetFolder -ItemType Directory > $null 2>&1
				#Write-Host "Ordner erstellt: $targetFolder"
			}

			# Use 7-Zip to extract the downloaded file into the target folder
			Write-Host "Die Datei wird extrahiert." -ForegroundColor "yellow"
			#& "$sevenZipPath" x "$downloadPath" -o"$targetFolder" -y
			& "$sevenZipPath" x "$downloadPath" -o"$targetFolder" -y *> $null

			#Write-Host "Die Datei wurde nach $targetFolder extrahiert." -ForegroundColor "yellow"
		} #else {
			#Write-Host "7-Zip ist nicht installiert. Bitte installieren Sie 7-Zip, um fortzufahren."
		#}
	}

	# Get the folder matching the pattern
	$matchingFolder = Get-ChildItem -Path $InstallationFolder -Directory | Where-Object {
		$_.Name -match $folderPattern
	} | Select-Object -First 1  # Get the first match if there are multiple

	if ($matchingFolder) {
		Write_LogEntry -Message "Passender Ordner gefunden: $($matchingFolder.FullName)" -Level "INFO"

		# Extract the version number
		#$localVersion = [regex]::Match($matchingFolder.Name, $versionPattern).Groups[1].Value
		$setupCfgPath = Join-Path -Path $matchingFolder.FullName -ChildPath "setup.cfg"
		
		# Check if setup.cfg exists
		if (Test-Path -Path $setupCfgPath) {
			# Read the content of setup.cfg
			$setupCfgContent = Get-Content -Path $setupCfgPath -Raw

			# Extract the version number from the <setup> tag
			$versionPattern = '<setup title="\$\{\{ProductTitle\}\}" version="([\d.]+)"'
			$localVersion = [regex]::Match($setupCfgContent, $versionPattern).Groups[1].Value

			Write_LogEntry -Message "Lokale Installationsdatei Version ermittelt: $($localVersion)" -Level "DEBUG"
		} else {
			Write_LogEntry -Message "setup.cfg nicht gefunden in: $($matchingFolder.FullName)" -Level "WARNING"
			Write-Host "setup.cfg file not found in the folder: $($matchingFolder.FullName)"
		}
		
		$result = GetVersionAndDownloadLink -CurrentLocalVersion $localVersion
		if ($result) {
			$downloadUrl = $($result.DownloadLink)
			$onlineVersion = $($result.VersionNumber)
			Write_LogEntry -Message "Online-Version ermittelt: $($onlineVersion); Download-URL ermittelt: $($downloadUrl)" -Level "DEBUG"

			#other url just gives a download link to the notebook driver
			#https://de.download.nvidia.com/Windows/566.36/566.36-desktop-win10-win11-64bit-international-nsd-dch-whql.exe
			$downloadUrl = "https://de.download.nvidia.com/Windows/$($onlineVersion)/$($onlineVersion)-desktop-win10-win11-64bit-international-nsd-dch-whql.exe"
			
			# Write logs for versions
			Write_LogEntry -Message "Lokale Version: $($localVersion); Online Version: $($onlineVersion)" -Level "INFO"
			
			# Check if both values are present
			if (-not $downloadUrl -or -not $onlineVersion) {
				Write_LogEntry -Message "Fehlende Download-URL oder Online-Version. Abbruch." -Level "ERROR"
				Write-Host "Einer oder beide Werte fehlen. Das Skript wird beendet." -ForegroundColor "red"
				exit
			}
		} else {
			Write_LogEntry -Message "Kein Treiber-Link online gefunden." -Level "WARNING"
			Write-Host "Kein Treiber-Link gefunden." -ForegroundColor "yellow"
			exit
		}
		
		Write-Host ""
		Write-Host "Lokale Version: $localVersion" -ForegroundColor "Cyan"
		Write-Host "Online Version: $onlineVersion" -ForegroundColor "Cyan"
		Write-Host ""
			
		if ([version]$localVersion -eq [version]$onlineVersion) {
			Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -ForegroundColor "DarkGray"
			Write_LogEntry -Message "$($ProgramName) ist aktuell. Keine Aktion erforderlich." -Level "INFO"
		}
		else {
			Write_LogEntry -Message "Online-Update verfügbar. Starte Download: $($downloadUrl)" -Level "INFO"

			# Extract the filename from the download link (last part of the URL)
			$downloadFileName = [System.IO.Path]::GetFileName($downloadUrl)

			# Download the newer version with the original filename
			$downloadPath = Join-Path $InstallationFolder $downloadFileName
			Write_LogEntry -Message "Download-Pfad: $($downloadPath)" -Level "DEBUG"
			
			Import-Module BitsTransfer
			$useBitTransfer = $null -ne (Get-Module -Name BitsTransfer -ListAvailable)

			if ($useBitTransfer) {
				Write_LogEntry -Message "Verwende BitsTransfer zum Herunterladen." -Level "DEBUG"
				Start-BitsTransfer -Source $downloadUrl -Destination $downloadPath
				Write_LogEntry -Message "Start-BitsTransfer abgeschlossen: $($downloadPath)" -Level "INFO"
			} else{
				Write_LogEntry -Message "Verwende WebClient zum Herunterladen." -Level "DEBUG"
				$webClient = New-Object System.Net.WebClient
				[void](Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath)
				$webClient.Dispose()
				Write_LogEntry -Message "WebClient-Download abgeschlossen: $($downloadPath)" -Level "INFO"
			}
				
			# Check if the file was completely downloaded
			if (Test-Path $downloadPath) {
				Write_LogEntry -Message "Download bestanden: $($downloadPath). Beginne Extraktion." -Level "INFO"
				ExtractExeFile
				Write_LogEntry -Message "Extraktion abgeschlossen für: $($downloadFileName)" -Level "SUCCESS"
				
				# Remove the old installer
				try {
					Remove-Item -Path $downloadPath -Force
					Remove-Item -Path $matchingFolder -Force -Recurse
					Write_LogEntry -Message "Alte Dateien entfernt: $($downloadPath) und $($matchingFolder.FullName)" -Level "INFO"
				} catch {
					Write_LogEntry -Message "Fehler beim Entfernen alter Dateien: $($_)" -Level "ERROR"
				}
				
				Write-Host ""
				Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor "green"
			} else {
				Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath)" -Level "ERROR"
				Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "red"
			}
		}
	} else {
		Write_LogEntry -Message "Kein passender Ordner in $($InstallationFolder) gefunden." -Level "ERROR"
		Write-Host "Kein passender Ordner in $InstallationFolder gefunden." -ForegroundColor "red"
	}

	Write-Host ""

	#Check Installed Version / Install if needed
	$matchingFolder = Get-ChildItem -Path $InstallationFolder -Directory | Where-Object {
		$_.Name -match $folderPattern
	} | Select-Object -First 1  # Get the first match if there are multiple

	#$localVersion = [regex]::Match($matchingFolder.Name, $versionPattern).Groups[1].Value
	$setupCfgPath = Join-Path -Path $matchingFolder.FullName -ChildPath "setup.cfg"

	# Check if setup.cfg exists
	if (Test-Path -Path $setupCfgPath) {
		# Read the content of setup.cfg
		$setupCfgContent = Get-Content -Path $setupCfgPath -Raw

		# Extract the version number from the <setup> tag
		$versionPattern = '<setup title="\$\{\{ProductTitle\}\}" version="([\d.]+)"'
		$localVersion = [regex]::Match($setupCfgContent, $versionPattern).Groups[1].Value

		Write_LogEntry -Message "Installationsdatei Version erneut ermittelt: $($localVersion)" -Level "DEBUG"
	} else {
		Write_LogEntry -Message "setup.cfg nicht gefunden beim zweiten Check: $($matchingFolder.FullName)" -Level "WARNING"
		Write-Host "setup.cfg file not found in the folder: $($matchingFolder.FullName)"
	}

	#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

	$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
	Write_LogEntry -Message "Prüfe Registry-Pfade: $($RegistryPaths -join '; ')" -Level "DEBUG"

	$Path = foreach ($RegPath in $RegistryPaths) {
		if (Test-Path $RegPath) {
			Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
		}
	}

	if ($null -ne $Path) {
		$installedVersion = $Path.DisplayVersion | Select-Object -First 1
		Write-Host "$ProgramName ist installiert." -ForegroundColor "green"
		Write-Host "	Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
		Write-Host "	Installationsdatei Version: $localVersion" -ForegroundColor "Cyan"
		Write_LogEntry -Message "Registry-Eintrag gefunden. Installierte Version: $($installedVersion). Installationsdatei Version: $($localVersion)" -Level "INFO"
		
		if ([version]$installedVersion -lt [version]$localVersion) {
			Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "magenta"
			Write_LogEntry -Message "Installierte Version ($($installedVersion)) älter als lokale Version ($($localVersion)). Update erforderlich." -Level "INFO"
			$Install = $true
		} elseif ([version]$installedVersion -eq [version]$localVersion) {
			Write-Host "		Installierte Version ist aktuell." -ForegroundColor "DarkGray"
			Write_LogEntry -Message "Installierte Version ist aktuell: $($installedVersion)" -Level "DEBUG"
			$Install = $false
		} else {
			#Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
			Write_LogEntry -Message "Installierte Version ($($installedVersion)) neuer als lokale Version ($($localVersion)). Keine Aktion." -Level "WARNING"
			$Install = $false
		}
	} else {
		Write_LogEntry -Message "Kein Registry-Eintrag für $($ProgramName) gefunden." -Level "WARNING"
		#Write-Host "$ProgramName is not installed on this system."
		$Install = $false
	}
	Write-Host ""

	#Install if needed
	if($InstallationFlag){
		Write_LogEntry -Message "InstallationFlag gesetzt -> Starte Installationsscript mit Flag." -Level "INFO"
		& $PSHostPath `
			-NoLogo -NoProfile -ExecutionPolicy Bypass `
			-File "$Serverip\Daten\Prog\InstallationScripts\Installation\TreiberAmdPcNvideaInstall.ps1" `
			-InstallationFlag
		Write_LogEntry -Message "Externes Installationsscript aufgerufen mit -InstallationFlag: $($PSHostPath) $($Serverip)\Daten\Prog\InstallationScripts\Installation\TreiberAmdPcNvideaInstall.ps1" -Level "DEBUG"
	}
	elseif($Install -eq $true){
		Write_LogEntry -Message "Update erforderlich -> Starte Installationsscript." -Level "INFO"
		& $PSHostPath `
			-NoLogo -NoProfile -ExecutionPolicy Bypass `
			-File "$Serverip\Daten\Prog\InstallationScripts\Installation\TreiberAmdPcNvideaInstall.ps1"
		Write_LogEntry -Message "Externes Installationsscript aufgerufen: $($PSHostPath) $($Serverip)\Daten\Prog\InstallationScripts\Installation\TreiberAmdPcNvideaInstall.ps1" -Level "DEBUG"
	}
	Write-Host ""
} else {
	Write-Host ""
	Write-Host "		Treiber sind NICHT für dieses System geeignet." -ForegroundColor "Blue"
	Write-Host ""
	Write_LogEntry -Message "System $($PCName) ist nicht Zielsystem für $($ProgramName). Keine Aktion." -Level "INFO"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
