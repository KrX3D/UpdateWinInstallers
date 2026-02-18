param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Java"
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
    Write_LogEntry -Message "Konfigurationsdatei gefunden und geladen: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

$InstallationFolder = "$InstallationFolder\Kicad"

$localInstallerPath = "$InstallationFolder\jdk*windows-x64_bin.msi"
Write_LogEntry -Message "Suche lokale Installer unter: $($localInstallerPath)" -Level "DEBUG"

# Get the local installer file
$localInstaller = Get-ChildItem -Path $localInstallerPath | Select-Object -First 1
Write_LogEntry -Message ("Gefundene lokale Installer-Datei: " + $([string]($localInstaller | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue))) -Level "DEBUG"

# Check if the local installer file exists
if ($localInstaller) {
    Write_LogEntry -Message "Lokaler Installer existiert: $($localInstaller.FullName)" -Level "INFO"

    # Extract the version number from the local installer file name
    $localVersionRegex = 'jdk-([\d._]+)_windows-x64_bin.msi'
    $localVersion = [regex]::Match($localInstaller.Name, $localVersionRegex).Groups[1].Value
    Write_LogEntry -Message "Lokale Version extrahiert aus Dateiname: $($localVersion)" -Level "DEBUG"

    # If the version contains an underscore, split it and take the second part as the version
    if ($localVersion -like '*_*') {
        $localVersion = ($localVersion -split '_')[1]
        Write_LogEntry -Message "Lokale Version nach Split angepasst: $($localVersion)" -Level "DEBUG"
    }

    # Retrieve the latest Java version from the official website
    $javaURL = "https://www.oracle.com/java/technologies/downloads/"  # Update the URL based on the latest version
    Write_LogEntry -Message "Rufe Oracle Java-Seite ab: $($javaURL)" -Level "INFO"
    $webContent = Invoke-RestMethod -Uri $javaURL -UseBasicParsing
    Write_LogEntry -Message "Abruf der Webseite abgeschlossen; Inhaltstyp: $($webContent.GetType().FullName)" -Level "DEBUG"

    # Extract the version number from the web content using a dynamic regex pattern
    #$latestVersionRegex = '<h3 id="java\d+">JDK Development Kit ([\d.]+) downloads<\/h3>'
    $latestVersionRegex = '<h3 id="java\d+">Java SE Development Kit ([\d.]+) downloads<\/h3>'
    $latestVersion = [regex]::Match($webContent, $latestVersionRegex).Groups[1].Value
    Write_LogEntry -Message "Extrahierte Online-Version aus Webseite: $($latestVersion)" -Level "DEBUG"
	
    # Check if the $latestVersion has fewer than 3 components (X.X.X format)
    $versionComponents = $latestVersion -split '\.'
    if ($versionComponents.Count -lt 3) {
        # Append zeros to make it X.X.X format
        while ($versionComponents.Count -lt 3) {
            $versionComponents += '0'
        }
        $latestVersion = $versionComponents -join '.'
        Write_LogEntry -Message "Online-Version nach Auffüllen: $($latestVersion)" -Level "DEBUG"
    }
	
	Write-Host ""
	Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
	Write-Host ""
	
    # Compare the local and latest versions
	if ([version]$latestVersion -gt [version]$localVersion) {
        Write_LogEntry -Message "Update verfügbar: Online $($latestVersion) > Lokal $($localVersion)" -Level "INFO"
        # Extract the download link using regular expressions			
		$downloadLinkRegex = 'href="(https:\/\/download\.oracle\.com\/java\/\d+\/latest\/jdk-([\d.]+)_windows-x64_bin\.msi)"'
		$match = [regex]::Match($webContent, $downloadLinkRegex)
		Write_LogEntry -Message "Versuche Download-Link mit Regex zu extrahieren." -Level "DEBUG"

		if ($match.Success) {
			$downloadLink = $match.Groups[1].Value
			$fileVersion = $match.Groups[2].Value
			Write_LogEntry -Message "Download-Link gefunden: $($downloadLink); FileVersion: $($fileVersion)" -Level "INFO"

			# Modify the download path to include the correct filename format
			$filename = "jdk-$fileVersion" + "_$latestVersion" + "_windows-x64_bin.msi"
			$downloadPath = Join-Path -Path $InstallationFolder -ChildPath $filename
			Write_LogEntry -Message "Zielpfad für Download bestimmt: $($downloadPath)" -Level "DEBUG"

			#Invoke-WebRequest -Uri $downloadLink -OutFile $downloadPath
			Write_LogEntry -Message "Starte Download von $($downloadLink) nach $($downloadPath)" -Level "INFO"
			$webClient = New-Object System.Net.WebClient
			$webClient.DownloadFile($downloadLink, $downloadPath)
			$webClient.Dispose()
			Write_LogEntry -Message "Download abgeschlossen; prüfe Existenz: $($downloadPath)" -Level "DEBUG"
		} else {
            Write_LogEntry -Message "Download-Link konnte nicht extrahiert werden (Regex-Match fehlgeschlagen)." -Level "ERROR"
        }
	
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
            Write_LogEntry -Message "Download erfolgreich: $($downloadPath). Entferne alten Installer: $($localInstaller.FullName)" -Level "INFO"
			# Remove the old installer
			Remove-Item -Path $localInstaller.FullName -Force
            Write_LogEntry -Message "Alter Installer entfernt: $($localInstaller.FullName)" -Level "DEBUG"

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
		} else {
            Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath) nicht vorhanden nach Abschluss." -Level "ERROR"
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
		}

        # You can perform any additional tasks here, such as installing the latest version
    } else {
        Write_LogEntry -Message "Kein Online Update verfügbar. $($ProgramName) ist aktuell (Online: $($latestVersion); Lokal: $($localVersion))." -Level "INFO"
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    }
} else {
    Write_LogEntry -Message "Lokaler Java-Installer nicht gefunden unter: $($localInstallerPath)" -Level "WARNING"
    #Write-Host "Local Java installer not found."
}

Write-Host ""
Write_LogEntry -Message "Starte Prüfung installierter Versionen (Registry)." -Level "DEBUG"

#Check Installed Version / Install if neded
$localInstaller = Get-ChildItem -Path $localInstallerPath | Select-Object -First 1
$localVersionRegex = 'jdk-([\d._]+)_windows-x64_bin.msi'
$localVersion = [regex]::Match($localInstaller.Name, $localVersionRegex).Groups[1].Value
if ($localVersion -like '*_*') {
	$localVersion = ($localVersion -split '_')[1]
    Write_LogEntry -Message "Lokale Version nach letztem Extrakt: $($localVersion)" -Level "DEBUG"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Zu prüfende Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
	$installedVersion = $Path.DisplayVersion | ForEach-Object {
		$versionParts = $_ -split '\.'
		if ($versionParts.Count -ge 3) {
			$versionParts[0..2] -join '.'
		} else {
			$_
		}
	} | Select-Object -First 1
    Write_LogEntry -Message "Gefundene installierte Version in Registry: $($installedVersion)" -Level "INFO"
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist älter als lokale Datei ($($localVersion)). Markiere Installation." -Level "INFO"
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write_LogEntry -Message "Installierte Version entspricht lokaler Datei: $($installedVersion)" -Level "DEBUG"
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$Install = $false
    } else {
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Datei ($($localVersion)). Kein Update nötig." -Level "WARNING"
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
    }
} else {
    Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden. Install-Flag auf $($false) gesetzt." -Level "INFO"
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1 mit Parameter -InstallationFlag" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Install Flag true: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1"
    Write_LogEntry -Message "Externes Installations-Skript für Java aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
    Finalize_LogSession | Out-Null
}
# === Ende Logger-Footer ===
