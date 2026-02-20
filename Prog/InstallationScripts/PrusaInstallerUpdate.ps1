param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "PrusaSlicer"
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
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

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
Write_LogEntry -Message "Berechneter Konfigurationspfad: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei geladen: $($configPath)" -Level "INFO"
} else {
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    exit
}

$InstallationFolder = "$InstallationFolder\3D"
Write_LogEntry -Message "InstallationFolder gesetzt: $($InstallationFolder)" -Level "DEBUG"
$InstallationFileFile = "$InstallationFolder\prusaslicer_*.exe"
Write_LogEntry -Message "Suchmuster für Installationsdateien: $($InstallationFileFile)" -Level "DEBUG"

$FoundFile = Get-ChildItem -Path $InstallationFileFile -ErrorAction SilentlyContinue | Select-Object -First 1
if ($FoundFile) {
    $InstallationFileName = $FoundFile.Name
    Write_LogEntry -Message "Lokale Installationsdatei gefunden: $($InstallationFileName)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei gefunden mit Muster: $($InstallationFileFile)" -Level "WARNING"
}

$localInstaller = "$InstallationFolder\$InstallationFileName"
try {
    $localVersion = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).VersionInfo.ProductVersion
    Write_LogEntry -Message "Lokale Installer-Version ermittelt: $($localVersion) aus Datei $($localInstaller)" -Level "DEBUG"
} catch {
    $localVersion = "0.0.0"
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Installer-Version für $($localInstaller): $($_)" -Level "WARNING"
}

# Define the GitHub API endpoint URL
$apiUrl = "https://api.github.com/repos/prusa3d/PrusaSlicer/releases/latest"
Write_LogEntry -Message "GitHub API URL: $($apiUrl)" -Level "DEBUG"

# --- HttpWebRequest (fast, low-level) with headers and robust error handling ---
# Ensure TLS 1.2
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    # ignore if not supported; continue
}

$headers = @{
    'User-Agent' = 'InstallationScripts/1.0'
    'Accept'     = 'application/vnd.github.v3+json'
}
if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
}

$release = $null
try {
    $request = [Net.WebRequest]::Create($apiUrl)
    $request.Method = "GET"
    $request.Timeout = 30000
    $request.UserAgent = $headers['User-Agent']
    $request.ContentType = "application/json"
    #$request.Headers.Add("Accept", $headers['Accept'])
	$request.Accept = $headers['Accept']  # Use property instead of Headers.Add()
    if ($headers.ContainsKey('Authorization')) {
        # Add Authorization header in a way that works with WebRequest
        $request.Headers.Add("Authorization", $headers['Authorization'])
    }
	
    Write_LogEntry -Message "Sende HttpWebRequest an GitHub API ($($apiUrl))" -Level "DEBUG"

    try {
        $response = $request.GetResponse()
        $stream = $response.GetResponseStream()
        $reader = [System.IO.StreamReader]::new($stream)
        $body = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()
        $response.Close()

        # Convert JSON body to object
        try {
            $release = $body | ConvertFrom-Json -ErrorAction Stop
            Write_LogEntry -Message "GitHub API Response empfangen; TagName: $($release.tag_name)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Parsen der GitHub-Antwort: $($_)" -Level "ERROR"
            $release = $null
        }
    } catch [System.Net.WebException] {
        $webEx = $_.Exception
        $errMsg = $webEx.Message
        # Try to read response body
        try {
            if ($webEx.Response -ne $null) {
                $errStream = $webEx.Response.GetResponseStream()
                $errReader = [System.IO.StreamReader]::new($errStream)
                $errBody = $errReader.ReadToEnd()
                $errReader.Close()
                $errStream.Close()
                # Try to parse JSON error message
                try {
                    $errJson = $errBody | ConvertFrom-Json -ErrorAction Stop
                    if ($errJson.message) { $errMsg = $errJson.message }
                } catch {
                    # leave errMsg as-is or use raw body if helpful
                    if ($errBody) { $errMsg = $errBody }
                }
            }
        } catch {
            # ignore read errors
        }

        if ($errMsg -and ($errMsg -match '(rate limit|rate_limit|rate limit exceeded|API rate limit|403)')) {
            Write_LogEntry -Message "GitHub API Rate-Limit / Zugriff verweigert erkannt: $($errMsg)" -Level "WARNING"
            Write_LogEntry -Message "Hinweis: Verwende ein GitHub-PAT (in PowerShellVariables.ps1 als \$GithubToken) oder env var to increase rate limits." -Level "DEBUG"
        } else {
            Write_LogEntry -Message "Fehler beim Abrufen der GitHub API: $($errMsg)" -Level "ERROR"
        }

        $release = $null
    }
} catch {
    Write_LogEntry -Message "Unerwarteter Fehler beim Aufbau des WebRequest: $($_)" -Level "ERROR"
    $release = $null
}
# --- end HttpWebRequest block ---

# If API failed, avoid crashing by setting $newVersion = $localVersion (no update available)
if (-not $release) {
    $newVersion = $localVersion
    Write_LogEntry -Message "Keine Online-Release-Information verfügbar; Online-Check übersprungen. Setze neue Version = lokale Version ($newVersion)" -Level "WARNING"
} else {
    # Parse release object as before
    try {
        $newVersion = $release.tag_name.TrimStart("version_")
        Write_LogEntry -Message "Extrahierte Online-Version: $($newVersion); Lokale Version: $($localVersion)" -Level "INFO"
    } catch {
        $newVersion = $localVersion
        Write_LogEntry -Message "Fehler beim Extrahieren der Online-Version; setze neue Version = lokale Version ($newVersion)" -Level "WARNING"
    }
}

Write-Host ""
Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
Write-Host "Online Version: $newVersion" -foregroundcolor "Cyan"
Write-Host ""

# initialize variable to avoid later undefined errors
$win64DownloadUrl = $null

if ([version]$newVersion -ne [version]$localVersion) {
    $newInstaller = "$InstallationFolder\prusaslicer_$newVersion.exe"
    Write_LogEntry -Message "Neue Version erkannt: $($newVersion). Neuer Installer-Pfad: $($newInstaller)" -Level "INFO"
	
	#Get Download Link
	$onlineVersionUrl = "https://help.prusa3d.com/downloads"
	Write_LogEntry -Message "Rufe Online-Version-Seite ab: $($onlineVersionUrl)" -Level "DEBUG"
	try {
		$onlineVersionHtml = Invoke-WebRequest -Uri $onlineVersionUrl -UseBasicParsing -ErrorAction Stop
		Write_LogEntry -Message "Online-Downloads-Seite abgerufen" -Level "DEBUG"
	} catch {
		Write_LogEntry -Message "Fehler beim Abrufen der Online-Downloads-Seite $($onlineVersionUrl): $($_)" -Level "ERROR"
		$onlineVersionHtml = $null
	}

	# Use the regex pattern to find download links for "win" and "standalone" (raw JSON-like snippets)
	$pattern = '\\"platform\\":\\"(win|standalone)\\",\\"show_linux_info\\":true,\\"file_url\\":\\"(https://[^"]+)\\"'
	$matchesPattern = @()
	if ($onlineVersionHtml) {
		$matchesPattern = [regex]::Matches($onlineVersionHtml.Content, $pattern)
		Write_LogEntry -Message "Anzahl gefundener Download-Matches auf der Seite: $($matchesPattern.Count)" -Level "DEBUG"
	} else {
		Write_LogEntry -Message "Kein HTML-Inhalt zum Parsen vorhanden." -Level "WARNING"
	}

	if ($matchesPattern.Count -gt 0) {
		$winVersions = @{}
		$standaloneVersions = @{}

		foreach ($match in $matchesPattern) {
			$platform = $match.Groups[1].Value
			$downloadLink = $match.Groups[2].Value

			# Normalize for checks
			$dlLower = $downloadLink.ToLower()

			# Skip clearly irrelevant items: firmware, archives, or anything not an exe
			if ($dlLower -match 'firmware' -or $dlLower -match '\.zip$' -or $dlLower -match '\.tar' -or $dlLower -match '\.gz$') {
				Write_LogEntry -Message "Skip non-driver asset (firmware/archive): $($downloadLink)" -Level "DEBUG"
				continue
			}
			# prefer only .exe files (after cleaning)
			if (-not ($downloadLink -match '\.exe')) {
				Write_LogEntry -Message "Skip non-exe asset: $($downloadLink)" -Level "DEBUG"
				continue
			}

			# Normalize the link ending to .exe (like you had before)
			$downloadLink = $downloadLink -replace '(\.exe.*)$', '.exe'

			# Heuristic: only accept driver / prusaslicer installer links for Windows
			# Accept if link path contains 'drivers' or filename contains typical windows installer markers
			$isDriverLike = ($dlLower -match '/downloads/drivers/') -or ($dlLower -match 'prusa3d_win') -or ($dlLower -match 'prusaslicer_win') -or ($dlLower -match 'prusa3d_win_') -or ($dlLower -match 'prusaslicer_win_standalone')

			# Platform-specific acceptance
			if ($platform -eq "win") {
				if (-not $isDriverLike) {
					Write_LogEntry -Message "Ignored WIN asset not driver-like: $($downloadLink)" -Level "DEBUG"
					continue
				}
			} elseif ($platform -eq "standalone") {
				# Some standalone assets are explicit: accept those containing 'standalone' or 'PrusaSlicer_Win_standalone'
				if (-not ($dlLower -match 'standalone' -or $dlLower -match 'prusaslicer_win_standalone')) {
					Write_LogEntry -Message "Ignored STANDALONE asset not standalone-like: $($downloadLink)" -Level "DEBUG"
					continue
				}
			}

			# Extract the version number using a regular expression (keep both formats)
			$versionMatch = [regex]::Match($downloadLink, '(\d+\.\d+\.\d+|\d+_\d+_\d+)')
			if ($versionMatch.Success) {
				$version = $versionMatch.Groups[1].Value -replace '_', '.'

				# Store the version link in the appropriate platform's hash table
				if ($platform -eq "win") {
					$winVersions[$version] = $downloadLink
					Write_LogEntry -Message "Win-Asset hinzugefügt: Version $($version); URL: $($downloadLink)" -Level "DEBUG"
				} elseif ($platform -eq "standalone") {
					$standaloneVersions[$version] = $downloadLink
					Write_LogEntry -Message "Standalone-Asset hinzugefügt: Version $($version); URL: $($downloadLink)" -Level "DEBUG"
				}
			} else {
				Write_LogEntry -Message "Version konnte nicht extrahiert aus Link, übersprungen: $($downloadLink)" -Level "DEBUG"
			}
		}

		# Get the newest version links for "win" and "standalone" if they exist
		$latestWinVersion = $null
		$latestStandaloneVersion = $null
		if ($winVersions.Keys.Count -gt 0) {
			$latestWinVersion = $winVersions.Keys | Sort-Object { [Version] $_ } | Select-Object -Last 1
		}
		if ($standaloneVersions.Keys.Count -gt 0) {
			$latestStandaloneVersion = $standaloneVersions.Keys | Sort-Object { [Version] $_ } | Select-Object -Last 1
		}

		Write_LogEntry -Message "Latest Win Version: $($latestWinVersion); Latest Standalone Version: $($latestStandaloneVersion)" -Level "DEBUG"

		# Choose final download URL with preference rules
		if ($latestWinVersion -and $latestStandaloneVersion) {
			if ([version]$latestWinVersion -eq [version]$latestStandaloneVersion) {
				$win64DownloadUrl = $standaloneVersions[$latestStandaloneVersion]
				Write_LogEntry -Message "Beide Versionen gleich; wähle Standalone URL: $($win64DownloadUrl)" -Level "DEBUG"
			} elseif ([version]$latestWinVersion -gt [version]$latestStandaloneVersion) {
				$win64DownloadUrl = $winVersions[$latestWinVersion]
				Write_LogEntry -Message "Wähle Win URL: $($win64DownloadUrl)" -Level "DEBUG"
			} else {
				$win64DownloadUrl = $standaloneVersions[$latestStandaloneVersion]
				Write_LogEntry -Message "Wähle Standalone URL: $($win64DownloadUrl)" -Level "DEBUG"
			}
		} elseif ($latestWinVersion) {
			$win64DownloadUrl = $winVersions[$latestWinVersion]
			Write_LogEntry -Message "Nur Win-Asset gefunden; wähle Win URL: $($win64DownloadUrl)" -Level "DEBUG"
		} elseif ($latestStandaloneVersion) {
			$win64DownloadUrl = $standaloneVersions[$latestStandaloneVersion]
			Write_LogEntry -Message "Nur Standalone-Asset gefunden; wähle Standalone URL: $($win64DownloadUrl)" -Level "DEBUG"
		} else {
			Write_LogEntry -Message "Keine passende Download-URL für Win/Standalone gefunden (nach Filter)." -Level "WARNING"
			$win64DownloadUrl = $null
		}
	} 

	# Compare if the download link is the same version as the newest Version from GIT
	if ($win64DownloadUrl) {
		# Extract the version number from the URL and convert underscores to periods
		$urlVersionMatch = [regex]::Match($win64DownloadUrl, '(\d+\.\d+\.\d+|\d+_\d+_\d+)')
		if ($urlVersionMatch.Success) {
			$urlVersion = $urlVersionMatch.Groups[1].Value -replace '_', '.'
			Write_LogEntry -Message "Version extrahiert aus Download-URL: $($urlVersion); Erwartete Online-Version: $($newVersion)" -Level "DEBUG"

			# Compare the extracted version with the desired version
			if ($newVersion -eq $urlVersion) {
				Write_LogEntry -Message "Versions-Check erfolgreich: $($newVersion) == $($urlVersion). Starte Download..." -Level "INFO"
				
				$webClient = New-Object System.Net.WebClient
	            try {
				    [void](Invoke-DownloadFile -Url $win64DownloadUrl -OutFile $newInstaller)
	                Write_LogEntry -Message "Download beendet: $($newInstaller)" -Level "DEBUG"
	            } catch {
	                Write_LogEntry -Message "Fehler beim Herunterladen von $($win64DownloadUrl): $($_)" -Level "ERROR"
	            } finally {
	                $webClient.Dispose()
	            }
				
				# Check if the file was completely downloaded
				if (Test-Path $newInstaller) {
					# Remove the old installer
					try {
						Remove-Item -Path $localInstaller -Force
						Write_LogEntry -Message "Alte Installer-Datei entfernt: $($localInstaller)" -Level "DEBUG"
					} catch {
						Write_LogEntry -Message "Fehler beim Entfernen der alten Installer-Datei $($localInstaller): $($_)" -Level "WARNING"
					}

					Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
	                Write_LogEntry -Message "$($ProgramName) Update erfolgreich: $($newInstaller)" -Level "SUCCESS"
				} else {
					Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
	                Write_LogEntry -Message "Download fehlgeschlagen; Datei nicht gefunden: $($newInstaller)" -Level "ERROR"
				}    
			} else {
				Write-Host "Versionen stimmen nicht überein. Gewünscht: $newVersion, URL: $urlVersion" -foregroundcolor "red"
	            Write_LogEntry -Message "Versionskonflikt: Gewünscht $($newVersion), URL-Version $($urlVersion)" -Level "WARNING"
			}
		} else {
			Write_LogEntry -Message "Version konnte nicht aus der Download-URL extrahiert werden: $($win64DownloadUrl)" -Level "WARNING"
		}
	} else {
		Write_LogEntry -Message "Kein geeigneter Download-URL gefunden; überspringe Download." -Level "WARNING"
	}
} else {
	Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    Write_LogEntry -Message "Kein Online Update verfügbar: Online $($newVersion) == Lokal $($localVersion)" -Level "INFO"
}

Write-Host ""

#Check Installed Version / Install if neded
$FoundFile = Get-ChildItem -Path $InstallationFileFile -ErrorAction SilentlyContinue | Select-Object -First 1
if ($FoundFile) {
    $InstallationFileName = $FoundFile.Name
    $localInstaller = "$InstallationFolder\$InstallationFileName"
    try {
        $localVersion = (Get-Item -LiteralPath $localInstaller -ErrorAction Stop).VersionInfo.ProductVersion
        Write_LogEntry -Message "Erneut lokale Installationsdatei ermittelt: $($localInstaller); Version: $($localVersion)" -Level "DEBUG"
    } catch {
        $localVersion = "0.0.0"
        Write_LogEntry -Message "Fehler beim Ermitteln der Version der Datei $($localInstaller): $($_)" -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei beim erneutem Check gefunden." -Level "WARNING"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registrypfade zur Suche: $($RegistryPaths -join '; ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$($ProgramName) in Registry gefunden; InstallierteVersion: $($installedVersion); InstallationsdateiVersion: $($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (Update erforderlich)" -Level "INFO"
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
    Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden (nicht installiert)." -Level "INFO"
}
Write-Host ""
Write_LogEntry -Message "Installationsprüfung abgeschlossen. Install variable: $($Install)" -Level "DEBUG"

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte Installationsskript mit Flag: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Installationsskript mit Flag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte Installationsskript (Update) ohne Flag: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1"
    Write_LogEntry -Message "Installationsskript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\PrusaInstallerInstallation.ps1" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===

