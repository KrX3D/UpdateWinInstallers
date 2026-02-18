param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinReducerEX100"
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

$InstallationFolder = "$NetworkShareDaten\Customize_Windows\Tools\WinReducerEX100"
$InstallationFileFile = "$InstallationFolder\WinReducerEX100_x64.exe"
$destinationPath = Join-Path $env:USERPROFILE "Desktop\WinReducerEX100"
$destinationFilePath = Join-Path $destinationPath "WinReducerEX100_x64.exe"

Write_LogEntry -Message "InstallationFolder: $($InstallationFolder)" -Level "DEBUG"
Write_LogEntry -Message "InstallationFileFile: $($InstallationFileFile)" -Level "DEBUG"
Write_LogEntry -Message "DestinationPath: $($destinationPath); DestinationFilePath: $($destinationFilePath)" -Level "DEBUG"

#$webpageURL = "https://www.winreducer.net/ex-series.html"  # Webpage URL to check for the version
$webpageURL = "https://www.winreducer.net/winreducer-ex-series.html"  # Webpage URL to check for the version
Write_LogEntry -Message "Hole Webseite: $($webpageURL)" -Level "INFO"

# Retrieve the webpage content
try {
    $webpageContent = (Invoke-WebRequest -Uri $webpageURL -UseBasicParsing).Content
    Write_LogEntry -Message "Webseite abgerufen; Inhalt Länge: $($webpageContent.Length)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen der Webseite $($webpageURL): $($_)" -Level "ERROR"
    $webpageContent = ""
}

# Extract the version using regular expressions
$versionRegex = '(?<=<div class="paragraph" style="text-align:center;"><font size="6"><strong><font color="#fff">)v(.*?)(?=<\/font>)'
$versionMatch = [regex]::Match($webpageContent, $versionRegex)
$versionText = $versionMatch.Groups[1].Value.Trim()
Write_LogEntry -Message "Extrahierte Online-Version (versionText): $($versionText)" -Level "DEBUG"

try {
    $localVersion = (Get-ItemProperty -Path $InstallationFileFile).VersionInfo.FileVersion
    Write_LogEntry -Message "Ermittelte lokale Version: $($localVersion) aus Datei $($InstallationFileFile)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Version von $($InstallationFileFile): $($_)" -Level "WARNING"
    $localVersion = ""
}

Write-Host ""
Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
Write-Host "Online Version: $versionText" -foregroundcolor "Cyan"
Write-Host ""
Write_LogEntry -Message "Lokale Version: $($localVersion); Online Version: $($versionText)" -Level "INFO"

# Compare the file versions
if ($versionText -ne '' -and ($versionText -gt  $localVersion)){		
    Write_LogEntry -Message "Online-Version ($($versionText)) ist größer als lokale Version ($($localVersion)); starte Updateprozess" -Level "INFO"

    # Search for the download link
    #$downloadURLRegex = '(?<=<a class="wsite-button wsite-button-large wsite-button-highlight" href=").*(?=winreducer_ex_series_x64.zip)'
	$downloadURLRegex = '<a class="wsite-button wsite-button-large wsite-button-highlight" href="(.+?)".*?>\s*<span class="wsite-button-inner">DOWNLOAD \(x64\)</span>'

    $downloadMatch = [regex]::Match($webpageContent, $downloadURLRegex)
    #$partialDownloadURL = $downloadMatch.Value
	
	#Extract the download URL from the capture group
    $partialDownloadURL = $downloadMatch.Groups[1].Value
    Write_LogEntry -Message "Gefundene partielle Download-URL: $($partialDownloadURL)" -Level "DEBUG"

    # Construct the complete download URL
    #$downloadURL = "https://www.winreducer.net$partialDownloadURL" + "winreducer_ex_series_x64.zip"
    $downloadURL = "https://www.winreducer.net$partialDownloadURL" 
    Write_LogEntry -Message "Konstruiere Download-URL: $($downloadURL)" -Level "DEBUG"

    # Set the download path to the Windows TEMP folder
    $downloadPath = Join-Path -Path $env:TEMP -ChildPath "WinReducerEX100_x64_new.zip"
    Write_LogEntry -Message "Downloadpfad gesetzt: $($downloadPath)" -Level "DEBUG"

    # Download the ZIP file
    #Invoke-WebRequest -Uri $downloadURL -OutFile $downloadPath
	$webClient = New-Object System.Net.WebClient
	
	#Kann gebraucht werden, wenn die Webseite eine Restriction hat.
	#Kann sein, dass man die Versionen ab und zu updaten muss
	$webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    try {
        Write_LogEntry -Message "Starte Download von $($downloadURL) nach $($downloadPath)" -Level "INFO"
        $webClient.DownloadFile($downloadURL, $downloadPath)
        Write_LogEntry -Message "Download von $($downloadURL) abgeschlossen nach $($downloadPath)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Download von $($downloadURL): $($_)" -Level "ERROR"
    } finally {
        $null = $webClient.Dispose()
    }
	
	# Check if the file was completely downloaded
	if (Test-Path $downloadPath) {
        Write_LogEntry -Message "Download vorhanden: $($downloadPath)" -Level "DEBUG"
		# Extract the desired folder from the ZIP file to the TEMP folder
        try {
            Expand-Archive -Path $downloadPath -DestinationPath $env:TEMP -Force
            Write_LogEntry -Message "Archiv $($downloadPath) entpackt nach $($env:TEMP)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Entpacken von $($downloadPath): $($_)" -Level "ERROR"
        }

		#Backup License and stuff from WinReducer Folder before replaceing it
        try {
            Move-Item -Path (Join-Path -Path $InstallationFolder -ChildPath "\HOME\SOFTWARE\x64") -Destination (Join-Path -Path $env:TEMP -ChildPath "WinReducer_EX_Series_x64\WinReducerEX100\HOME\SOFTWARE") -Force
            Write_LogEntry -Message "Backup verschoben: HOME\x64 -> $($env:TEMP)\WinReducer_EX_Series_x64\WinReducerEX100\HOME\SOFTWARE" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Verschieben des HOME\x64 Backups: $($_)" -Level "WARNING"
        }
        try {
            Move-Item -Path (Join-Path -Path $InstallationFolder -ChildPath "\HOME\SOFTWARE\WinReducerEX100.xml") -Destination (Join-Path -Path $env:TEMP -ChildPath "WinReducer_EX_Series_x64\WinReducerEX100\HOME\SOFTWARE") -Force
            Write_LogEntry -Message "Backup verschoben: WinReducerEX100.xml -> $($env:TEMP)\WinReducer_EX_Series_x64\WinReducerEX100\HOME\SOFTWARE" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Verschieben der WinReducerEX100.xml: $($_)" -Level "WARNING"
        }
		
		# Move the desired folder to the download folderInstallationFolder
        try {
             Remove-Item -Path $InstallationFolder -Recurse -Force 
             Write_LogEntry -Message "Altes Installationsverzeichnis entfernt: $($InstallationFolder)" -Level "DEBUG"
        } catch {
             Write_LogEntry -Message "Fehler beim Entfernen des Installationsverzeichnisses $($InstallationFolder): $($_)" -Level "WARNING"
        }
		$extractedFolderPath = Join-Path -Path $env:TEMP -ChildPath "WinReducer_EX_Series_x64\WinReducerEX100"
        try {
            Move-Item -Path $extractedFolderPath -Destination (Split-Path -Path $InstallationFolder -Parent) -Force
            Write_LogEntry -Message "Extrahierter Ordner $($extractedFolderPath) nach $((Split-Path -Path $InstallationFolder -Parent)) verschoben" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Verschieben des extrahierten Ordners $($extractedFolderPath): $($_)" -Level "ERROR"
        }
		
		# Optionally, you can replace the existing folder with the moved one
		$extractedFolderPath = Join-Path -Path $env:TEMP -ChildPath "WinReducer_EX_Series_x64"
        try {
            Remove-Item -Path $extractedFolderPath -Recurse -Force
            Write_LogEntry -Message "Temporärer extrahierter Ordner entfernt: $($extractedFolderPath)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Entfernen des temporären extrahierten Ordners $($extractedFolderPath): $($_)" -Level "WARNING"
        }
        try {
            Remove-Item -Path $downloadPath -Recurse -Force 
            Write_LogEntry -Message "Archiv entfernt: $($downloadPath)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Entfernen des Archivs $($downloadPath): $($_)" -Level "WARNING"
        }

		Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
        Write_LogEntry -Message "$($ProgramName) wurde aktualisiert" -Level "SUCCESS"
	} else {
		Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
        Write_LogEntry -Message "Download nicht vorhanden: $($downloadPath). Update fehlgeschlagen." -Level "ERROR"
	}
} else {
	Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    Write_LogEntry -Message "Kein Online-Update verfügbar: Online $($versionText) <= Local $($localVersion) oder keine Online-Version ermittelt" -Level "INFO"
}

Write-Host ""
Write_LogEntry -Message "Prüfung auf Installation/Installationsbedarf beginnt" -Level "DEBUG"

#Check Installed Version / Install if neded
try {
    $localVersion = (Get-ItemProperty -Path $InstallationFileFile).VersionInfo.FileVersion
    Write_LogEntry -Message "Erneut ermittelte lokale Dateiversion: $($localVersion)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim erneuten Ermitteln der lokalen Dateiversion $($InstallationFileFile): $($_)" -Level "WARNING"
    $localVersion = ""
}

if (Test-Path $destinationFilePath) {
    try {
        $installedVersion = (Get-ItemProperty -Path $destinationFilePath).VersionInfo.FileVersion
        Write_LogEntry -Message "Gefundene installierte Version am Zielpfad $($destinationFilePath): $($installedVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der installierten Version von $($destinationFilePath): $($_)" -Level "WARNING"
        $installedVersion = $null
    }
} else {
    Write_LogEntry -Message "Ziel-Dateipfad existiert nicht: $($destinationFilePath)" -Level "DEBUG"
}

if ($null -ne $installedVersion) {
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Installationsstatus: installiert; InstalledVersion: $($installedVersion); LocalVersion: $($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Install erforderlich: Installed $($installedVersion) < Local $($localVersion)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$Install = $false
        Write_LogEntry -Message "Install nicht erforderlich: InstalledVersion == LocalVersion ($($localVersion))" -Level "INFO"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
        Write_LogEntry -Message "Install nicht ausgeführt: InstalledVersion ($($installedVersion)) > LocalVersion ($($localVersion))" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
    Write_LogEntry -Message "Keine installierte Version gefunden; Install=false" -Level "DEBUG"
}
Write-Host ""

#Install if needed
if($Install -eq $true -or $InstallationFlag)
{
	Write-Host "WinReducerEX100 wird kopert"
    Write_LogEntry -Message "Starte Kopiervorgang: Quelle $($InstallationFileFile) -> Ziel $($destinationPath) (Install=$($Install); InstallationFlag=$($InstallationFlag))" -Level "INFO"

	if (Test-Path $InstallationFileFile) {
        try {
		    Copy-Item -Path $InstallationFolder -Destination $destinationPath -Recurse -Force
            Write_LogEntry -Message "Kopiervorgang erfolgreich: $($InstallationFolder) -> $($destinationPath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Kopieren $($InstallationFolder) nach $($destinationPath): $($_)" -Level "ERROR"
        }
	} else {
        Write_LogEntry -Message "Quellpfad zum Kopieren existiert nicht: $($InstallationFileFile)" -Level "ERROR"
    }
}
Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht" -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===
