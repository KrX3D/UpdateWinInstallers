param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "PortUp"
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

$InstallationFolder = "$NetworkShareDaten\Customize_Windows"
Write_LogEntry -Message "InstallationFolder gesetzt: $($InstallationFolder)" -Level "DEBUG"

$localFilePath = "$InstallationFolder\Tools\PortableUpdate\PortUp.exe"
$destinationPath = Join-Path $env:USERPROFILE "Desktop\Reducer\PortUp.exe"
Write_LogEntry -Message "Lokaler Datei-Pfad: $($localFilePath); Ziel-Pfad: $($destinationPath)" -Level "DEBUG"

$webPageUrl = "https://www.portableupdate.com/download"
Write_LogEntry -Message "Webseite für Versionsprüfung: $($webPageUrl)" -Level "DEBUG"

# Get the local file version
$localVersion = (Get-ItemProperty -Path $localFilePath).VersionInfo.FileVersion
Write_LogEntry -Message "Lokale Dateiversion aus $($localFilePath) ermittelt: $($localVersion)" -Level "DEBUG"

$versionParts = $localVersion.Split('.')
$trimmedVersion = ($versionParts[0], $versionParts[1], $versionParts[2] -replace '^0+(\d)', '$1') -join '.'
$localVersion = $trimmedVersion
Write_LogEntry -Message "Bereinigte lokale Version: $($localVersion)" -Level "DEBUG"

# Retrieve the web page content
$webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing
Write_LogEntry -Message "Webseite abgerufen: $($webPageUrl); InhaltLaenge: $($webPageContent.Content.Length)" -Level "DEBUG"

$versionPattern = 'Portable Update (\d+\.\d+\.\d+)'
$versionMatch = [regex]::Match($webPageContent.Content, $versionPattern)
Write_LogEntry -Message "VersionPattern: $($versionPattern); MatchSuccess: $($versionMatch.Success)" -Level "DEBUG"

if ($versionMatch.Success) {
    $onlineversion = $versionMatch.Groups[1].Value
    $downloadLink = "https://file.portableupdate.com/downloads/PortUp_$onlineversion.zip"
	$downloadFileName = Split-Path -Path $downloadLink -Leaf
	$tempFilePath = Join-Path -Path $env:TEMP -ChildPath $downloadFileName

	Write-Host ""
	Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $onlineversion" -foregroundcolor "Cyan"
	Write-Host ""
	Write_LogEntry -Message "Onlineversion ermittelt: $($onlineversion); DownloadLink: $($downloadLink)" -Level "INFO"
	
	# Compare the local and remote file versions
	if ($onlineversion -gt $localVersion) {
        Write_LogEntry -Message "Update verfügbar: Online ($($onlineversion)) > Lokal ($($localVersion))" -Level "INFO"
        # Download the updated installer to the temporary folder
        #Invoke-WebRequest -Uri $downloadLink -OutFile $tempFilePath
        $webClient = New-Object System.Net.WebClient
        try {
            Write_LogEntry -Message "Starte Download von $($downloadLink) nach $($tempFilePath)" -Level "INFO"
            $webClient.DownloadFile($downloadLink, $tempFilePath)
            Write_LogEntry -Message "Download abgeschlossen: $($tempFilePath)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Herunterladen von $($downloadLink): $($($_.Exception.Message))" -Level "ERROR"
        } finally {
            $webClient.Dispose()
        }
		
		# Check if the file was completely downloaded
		if (Test-Path $tempFilePath) {
			try {
				# Remove the old installer
				Remove-Item -Path $localFilePath -Force
				Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFilePath)" -Level "DEBUG"
			} catch {
				Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei $($localFilePath): $($($_.Exception.Message))" -Level "WARNING"
			}
			
			# Extract the downloaded file to the installation folder
			try {
				Expand-Archive -Path $tempFilePath -DestinationPath $InstallationFolder -Force
				Write_LogEntry -Message "Archiv entpackt von $($tempFilePath) nach $($InstallationFolder)" -Level "SUCCESS"
			} catch {
				Write_LogEntry -Message "Fehler beim Entpacken des Archivs $($tempFilePath): $($($_.Exception.Message))" -Level "ERROR"
			}

			# Remove the temporary file
			try {
				Remove-Item $tempFilePath -Force
				Write_LogEntry -Message "Temporäre Datei entfernt: $($tempFilePath)" -Level "DEBUG"
			} catch {
				Write_LogEntry -Message "Fehler beim Entfernen der temporären Datei $($tempFilePath): $($($_.Exception.Message))" -Level "WARNING"
			}

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) Update erfolgreich auf Version $($onlineversion)" -Level "SUCCESS"
		} else {
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
            Write_LogEntry -Message "Download fehlgeschlagen; Datei nicht gefunden: $($tempFilePath)" -Level "ERROR"
		}
	} else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Online Update verfügbar. Online: $($onlineversion); Lokal: $($localVersion)" -Level "INFO"
	}
} else {
    Write_LogEntry -Message "Download link / Onlineversion nicht gefunden auf der Webseite: $($webPageUrl)" -Level "WARNING"
    #Write-Host "Download link not found on the website."
}

Write-Host ""

#Check Installed Version / Install if neded
$localVersion = (Get-ItemProperty -Path $localFilePath).VersionInfo.FileVersion
Write_LogEntry -Message "Erneut lokale Version von $($localFilePath) ermittelt: $($localVersion)" -Level "DEBUG"

if (Test-Path $destinationPath) {
	$installedVersion = (Get-ItemProperty -Path $destinationPath).VersionInfo.FileVersion
	Write_LogEntry -Message "Gefundene installierte Version am Zielpfad $($destinationPath): $($installedVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Zielpfad nicht gefunden: $($destinationPath)" -Level "DEBUG"
}

if ($null -ne $installedVersion) {
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	Write_LogEntry -Message "$($ProgramName) ist installiert; InstallierteVersion: $($installedVersion); InstallationsdateiVersion: $($localVersion)" -Level "INFO"
	
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
    Write_LogEntry -Message "$($ProgramName) nicht installiert am Zielpfad: $($destinationPath)" -Level "INFO"
}
Write-Host ""
Write_LogEntry -Message "Installationsprüfung abgeschlossen. Install variable: $($Install); InstallationFlag: $($InstallationFlag)" -Level "DEBUG"

#Install if needed
if($Install -eq $true -or $InstallationFlag)
{
	Write-Host "PortUp wird kopert"
    Write_LogEntry -Message "Starte Kopiervorgang: Quelle $($localFilePath) -> Ziel $($destinationPath)" -Level "INFO"

	if (Test-Path $localFilePath) {
		Copy-Item -Path $localFilePath -Destination $destinationPath -Force
        Write_LogEntry -Message "Datei kopiert: $($localFilePath) -> $($destinationPath)" -Level "SUCCESS"
	} else {
        Write_LogEntry -Message "Quell-Datei nicht gefunden, Kopiervorgang abgebrochen: $($localFilePath)" -Level "ERROR"
	}
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
