param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "EstlCam"
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
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PSScriptRoot: $($PSScriptRoot)" -Level "DEBUG"

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Berechneter Konfigurationspfad: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei $($configPath) gefunden und importiert." -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Script beendet wegen fehlender Konfigurationsdatei: $($configPath)" -Level "ERROR"
    Finalize_LogSession
    exit
}

$InstallationFolder = "$InstallationFolder\CNC"
Write_LogEntry -Message "InstallationFolder gesetzt auf: $($InstallationFolder)" -Level "DEBUG"

$filenamePattern = "Estlcam_64_*.exe"

function Get-EstlcamBuildFromFilename {
    param([Parameter(Mandatory)][string]$Name)
    # Match: Estlcam_64_13000.exe
    $m = [regex]::Match($Name, 'Estlcam_64_(\d+)\.exe', 'IgnoreCase')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return $null
}

function Get-LatestOnlineBuild {
    param([Parameter(Mandatory)][string]$Url)

    try {
        Write_LogEntry -Message "Rufe Online-Version ab: $Url" -Level "INFO"
        $webContent = Invoke-WebRequest -Uri $Url -UseBasicParsing -ErrorAction Stop
        Write_LogEntry -Message "Online-Seite abgerufen: $Url; Inhalt-Länge: $($webContent.Content.Length)" -Level "DEBUG"

        # Extract ALL builds from content, take max
        $matches = [regex]::Matches($webContent.Content, 'Estlcam_64_(\d+)\.exe', 'IgnoreCase')
        if ($matches.Count -eq 0) {
            Write_LogEntry -Message "Konnte keine Estlcam_64_####.exe im HTML finden." -Level "ERROR"
            return @{ Build = $null; Web = $webContent }
        }

        $builds = $matches | ForEach-Object { [int]$_.Groups[1].Value }
        $maxBuild = ($builds | Measure-Object -Maximum).Maximum

        return @{ Build = $maxBuild; Web = $webContent }
    }
    catch {
        Write_LogEntry -Message "Fehler beim Abruf der Online-Version: $($_.Exception.Message)" -Level "ERROR"
        return @{ Build = $null; Web = $null }
    }
}

# Get latest local installer
$latestFile = Get-ChildItem -Path $InstallationFolder -Filter $filenamePattern -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

Write_LogEntry -Message ("Ermittelte letzte lokale Datei: " + $([string]($latestFile | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue))) -Level "DEBUG"

if ($latestFile) {
    $localVersion = Get-EstlcamBuildFromFilename -Name $latestFile.Name
    if ($null -eq $localVersion) {
        Write_LogEntry -Message "Lokale Datei passt nicht zum erwarteten Muster: $($latestFile.Name)" -Level "ERROR"
    }

    Write_LogEntry -Message "Lokale Datei gefunden: $($latestFile.Name); Lokale Build-Version: $localVersion" -Level "INFO"

    # Check online for a newer version
    $latestVersionUrl = "http://www.estlcam.de/download.htm"  # URL to check for the latest version
    $online = Get-LatestOnlineBuild -Url $latestVersionUrl
    $latestVersion = $online.Build
    $webContent = $online.Web

    if ($null -eq $latestVersion) {
        Write-Host ""
        Write-Host "Online-Version konnte nicht ermittelt werden." -ForegroundColor "Red"
    }

    Write_LogEntry -Message "Online Build-Version extrahiert: $latestVersion" -Level "INFO"

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -ForegroundColor "Cyan"
    Write-Host "Online Version: $latestVersion" -ForegroundColor "Cyan"
    Write-Host ""

    if ($latestVersion -gt $localVersion) {
        Write_LogEntry -Message "Online-Version ($($latestVersion)) ist neuer als lokale Version ($($localVersion)). Suche Download-Link..." -Level "INFO"
        # Find the download link for the newer version
        # Prefer using parsed links first (if present), fallback to direct downloads path.
        $downloadLink = $null
        if ($webContent -and $webContent.Links) {
            $downloadLink = ($webContent.Links | Where-Object { $_.href -like "*Estlcam_64_$latestVersion.exe" } | Select-Object -First 1).href
        }

        if ($downloadLink) {
	        $downloadUrl  = "http://www.estlcam.de/" + $downloadLink.TrimStart('/')
            $downloadPath = Join-Path -Path $InstallationFolder -ChildPath "Estlcam_64_$latestVersion.exe"
            Write_LogEntry -Message "Download-Link gefunden: $($downloadUrl); Zielpfad: $($downloadPath)" -Level "DEBUG"

	        try {
				#Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath
				Write_LogEntry -Message "Starte Download von $($downloadUrl) nach $($downloadPath)" -Level "INFO"
				$webClient = New-Object System.Net.WebClient
	            $webClient.DownloadFile($downloadUrl, $downloadPath)
	            $webClient.Dispose()
	            Write_LogEntry -Message "Download abgeschlossen." -Level "SUCCESS"
	        }
	        catch {
	            Write_LogEntry -Message "Download fehlgeschlagen: $($_.Exception.Message)" -Level "ERROR"
	            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "Red"
	        }

			# Check if the file was completely downloaded
			if (Test-Path $downloadPath) {
                Write_LogEntry -Message "Download erfolgreich: $($downloadPath). Entferne alte Datei: $($latestFile.FullName)" -Level "INFO"
				# Remove the old installer
                try {
                    Remove-Item -Path $latestFile.FullName -Force
                    Write_LogEntry -Message "Alter Installer entfernt: $($latestFile.FullName)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message "Konnte alten Installer nicht löschen: $($latestFile.FullName) - $($_.Exception.Message)" -Level "WARNING"
                }

				Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
			} else {
                Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath) wurde nicht gefunden." -Level "ERROR"
				Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
			}
        } else {
            Write_LogEntry -Message "Kein Download-Link für Version $($latestVersion) gefunden." -Level "WARNING"
            #Write-Host "Download link not found for the newer version."
        }
		
    } else {
        Write_LogEntry -Message "Kein Online Update verfügbar. $($ProgramName) ist aktuell (Online: $($latestVersion), Lokal: $($localVersion))." -Level "INFO"
		Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -foregroundcolor "DarkGray"
    }
} else {
    Write_LogEntry -Message "Keine passende lokale Datei im Verzeichnis $($InstallationFolder) gefunden (Pattern: $($filenamePattern))." -Level "WARNING"
    #Write-Host "No matching file found in the specified directory."
}

Write-Host ""
Write_LogEntry -Message "Erneute lokale Datei-Überprüfung vor Installation/Registrierung." -Level "DEBUG"

#Check Installed Version / Install if needed
$FoundFile = Get-ChildItem -Path $InstallationFolder -Filter $filenamePattern -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

Write_LogEntry -Message ("Gefundene Datei für Installation: " + $([string]($FoundFile | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue))) -Level "DEBUG"


$localVersion = Get-EstlcamBuildFromFilename -Name $FoundFile.Name
Write_LogEntry -Message "Lokale Installationsdatei: $($InstallationFileName); Version: $($localVersion)" -Level "DEBUG"

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Prüfung: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
    # DisplayVersion may be string; force int where possible
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Gefundene installierte Version von $($ProgramName): $($installedVersion); Installationsdatei Version: $($localVersion)" -Level "INFO"

    if ($installedVersion -lt $localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist älter als lokale Installationsdatei ($($localVersion)). Update wird markiert." -Level "INFO"
		$Install = $true
    } elseif ($installedVersion -eq $localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Installierte Version ist aktuell: $($installedVersion)" -Level "DEBUG"
		$Install = $false
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Installationsdatei ($($localVersion)). Kein Update nötig." -Level "WARNING"
		$Install = $false
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
    Write_LogEntry -Message "$($ProgramName) nicht in der Registry gefunden. Setze Install-Flag auf $($false)." -Level "INFO"
	$Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\EstlCamInstallation.ps1 mit Parameter -InstallationFlag" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\EstlCamInstallation.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen." -Level "DEBUG"
}
elseif($Install -eq $true){
    Write_LogEntry -Message "Install-Flag true: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\EstlCamInstallation.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\EstlCamInstallation.ps1"
    Write_LogEntry -Message "Externes Installations-Skript ohne zusätzliche Parameter aufgerufen." -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
    Finalize_LogSession | Out-Null
}
# === Ende Logger-Footer ===
