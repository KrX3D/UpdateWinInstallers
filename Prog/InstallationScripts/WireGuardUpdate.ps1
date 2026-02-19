param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WireGuard"
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

# Define the directory path and file wildcard
$InstallationFolder = "$InstallationFolder\WireGuard"
Write_LogEntry -Message "InstallationFolder gesetzt auf: $($InstallationFolder)" -Level "DEBUG"

$ProgramName = "WireGuard"

$installerPath = "$InstallationFolder\wireguard-amd64-*.msi"
Write_LogEntry -Message "Installer-Pfad (Wildcard): $($installerPath)" -Level "DEBUG"
$installerFile = Get-InstallerFilePath -PathPattern $installerPath

# Check if the installer file exists
if ($installerFile) {
    Write_LogEntry -Message "Gefundene Installationsdatei: $($installerFile.FullName)" -Level "INFO"
    # Extract the version number from the file name
    $versionPattern = 'wireguard-amd64-(\d+\.\d+\.\d+)\.msi'
    $localVersion = Get-InstallerFileVersion -FilePath $installerFile.FullName -FileNameRegex $versionPattern -Source FileName
    Write_LogEntry -Message "Lokale Installationsdatei Version ermittelt: $($localVersion)" -Level "DEBUG"
    
    # Retrieve the latest version online from the GitHub repository tags
    $repositoryUrl = "https://github.com/WireGuard/wireguard-windows/tags"
    Write_LogEntry -Message "Hole Repository-Tags von: $($repositoryUrl)" -Level "INFO"
    $webPageContent = Invoke-RestMethod -Uri $repositoryUrl -UseBasicParsing
    Write_LogEntry -Message "Repository-Tags abgerufen; Länge Content: $($webPageContent.Length)" -Level "DEBUG"
    $versionPattern = "/WireGuard/wireguard-windows/releases/tag/v(\d+\.\d+\.\d+)"
    $latestVersion = [regex]::Matches($webPageContent, $versionPattern) |
        ForEach-Object { $_.Groups[1].Value } |
        Sort-Object -Descending |
        Select-Object -First 1
    Write_LogEntry -Message "Ermittelte Online-Version: $($latestVersion)" -Level "DEBUG"

	Write-Host ""
	Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
	Write-Host ""
	Write_LogEntry -Message "Vergleich Local: $($localVersion) vs Online: $($latestVersion)" -Level "INFO"
	
    # Compare the installed version with the latest version
    if ($localVersion -lt $latestVersion) {
        Write_LogEntry -Message "Online-Version neuer als lokal: $($latestVersion) > $($localVersion) - Download wird gestartet" -Level "INFO"
        # Construct the download URL for the newer version
        $downloadUrl = "https://download.wireguard.com/windows-client/wireguard-amd64-$latestVersion.msi"
        Write_LogEntry -Message "Download-URL konstruiert: $($downloadUrl)" -Level "DEBUG"

        # Set the download path for the newer version
        $downloadPath = "$InstallationFolder\wireguard-amd64-$latestVersion.msi"
        Write_LogEntry -Message "Download-Pfad gesetzt: $($downloadPath)" -Level "DEBUG"

        # Download the newer version
        #Invoke-RestMethod -Uri $downloadUrl -OutFile $downloadPath
        $webClient = New-Object System.Net.WebClient
        try {
            [void](Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath)
            Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Herunterladen $($downloadUrl) nach $($downloadPath): $($_)" -Level "ERROR"
        } finally {
            $null = $webClient.Dispose()
        }
		
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
			# Remove the old installer
            try {
			    Remove-Item -Path $installerFile.FullName -Force
                Write_LogEntry -Message "Alte Installationsdatei entfernt: $($installerFile.FullName)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei $($installerFile.FullName): $($_)" -Level "WARNING"
            }

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) Update erfolgreich: $($downloadPath)" -Level "SUCCESS"
		} else {
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
            Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath) nicht gefunden nach Download" -Level "ERROR"
		}
    } else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Update verfügbar: Online $($latestVersion) <= Local $($localVersion)" -Level "INFO"
    }
} else {
    Write_LogEntry -Message "Keine WireGuard-Installationsdatei gefunden im Pfad: $($installerPath)" -Level "WARNING"
    #Write-Host "No WireGuard installer found in the specified path."
}

Write-Host ""

#Check Installed Version / Install if neded
$installerFile = Get-InstallerFilePath -PathPattern $installerPath
if ($installerFile) {
    $versionPattern = 'wireguard-amd64-(\d+\.\d+\.\d+)\.msi'
    $localVersion = Get-InstallerFileVersion -FilePath $installerFile.FullName -FileNameRegex $versionPattern -Source FileName
    Write_LogEntry -Message "Ermittelte lokale Dateiversion nach erneutem Scan: $($localVersion) (Datei: $($installerFile.FullName))" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine Installationsdatei gefunden beim zweiten Scan: $($installerPath)" -Level "DEBUG"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Suche konfiguriert: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registry-Pfad existiert: $($RegPath)" -Level "DEBUG"
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
    Write_LogEntry -Message "Gefundene installierte Version aus Registry: $($installedVersion); Datei-Version: $($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Install erforderlich: Registry $($installedVersion) < Datei $($localVersion)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$Install = $false
        Write_LogEntry -Message "Install nicht erforderlich: InstalledVersion == LocalVersion ($($localVersion))" -Level "INFO"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
        Write_LogEntry -Message "InstalledVersion ($($installedVersion)) > LocalVersion ($($localVersion)); keine Aktion" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
    Write_LogEntry -Message "Keine Registry-Einträge für $($ProgramName) gefunden" -Level "DEBUG"
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "Starte externes Installationsskript aufgrund InstallationFlag" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\WireGuardInstall.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WireGuardInstall.ps1 mit -InstallationFlag" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte externes Installationsskript aufgrund Install=true" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\WireGuardInstall.ps1"
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WireGuardInstall.ps1" -Level "DEBUG"
}
Write-Host ""

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===
