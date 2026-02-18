param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Oracle VirtualBox"
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

$InstallationFileVirtualBox = "$InstallationFolder\VirtualBox-*.exe"
$InstallationFileVirtualBoxExtension = "$InstallationFolder\Oracle_VirtualBox_Extension_Pack-*.vbox-extpack"
Write_LogEntry -Message "Installationsordner Muster: $($InstallationFileVirtualBox); Extension Pack Muster: $($InstallationFileVirtualBoxExtension)" -Level "DEBUG"

function Get-LatestVersion {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    Write_LogEntry -Message "Hole LATEST.TXT von $($Url)" -Level "DEBUG"
    (Invoke-WebRequest -Uri $Url -UseBasicParsing).Content.Trim()
}

function DownloadFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Write_LogEntry -Message "Starte Download von $($Url) nach $($Path)" -Level "INFO"
    #Invoke-WebRequest -Uri $Url -OutFile $Path
	$webClient = New-Object System.Net.WebClient
	try {
		$webClient.DownloadFile($Url, $Path)
		Write_LogEntry -Message "Download erfolgreich: $($Path)" -Level "SUCCESS"
	} catch {
		Write_LogEntry -Message "Fehler beim Download $($Url) -> $($Path): $($_)" -Level "ERROR"
	} finally {
		$webClient.Dispose()
	}
}

function Compare-VersionNumbers {
    param (
        [string]$existingVersion,
        [string]$fullVersionNumber
    )

    # Split the version numbers into their components
    $existingVersionComponents = $existingVersion -split '\.'
    $fullVersionComponents = $fullVersionNumber -split '\.'

    Write_LogEntry -Message "Vergleiche Versionen: Lokal=$($existingVersion); Remote=$($fullVersionNumber)" -Level "DEBUG"

    # Determine the number of components to compare
    $minComponents = [math]::Min($existingVersionComponents.Length, $fullVersionComponents.Length)

    # Loop through the components and compare them
    for ($i = 0; $i -lt $minComponents; $i++) {
        $existingComponent = [int]$existingVersionComponents[$i]
        $fullComponent = [int]$fullVersionComponents[$i]

        if ($existingComponent -lt $fullComponent) {
            Write_LogEntry -Message "Version $($fullVersionNumber) ist neuer als $($existingVersion) (Komponente $i)" -Level "INFO"
            return $true
        } elseif ($existingComponent -gt $fullComponent) {
            Write_LogEntry -Message "Version $($existingVersion) ist neuer als $($fullVersionNumber) (Komponente $i)" -Level "DEBUG"
            return $false
        }
    }

    # If all components are equal, check if one version has more components
    if ($existingVersionComponents.Length -lt $fullVersionComponents.Length) {
        Write_LogEntry -Message "Remote-Version hat mehr Komponenten -> Remote ist neuer" -Level "INFO"
        return $true
    } elseif ($existingVersionComponents.Length -gt $fullVersionComponents.Length) {
        Write_LogEntry -Message "Lokale Version hat mehr Komponenten -> Lokal ist neuer" -Level "DEBUG"
        return $false
    } else {
        Write_LogEntry -Message "Versionen gleich: $($existingVersion) == $($fullVersionNumber)" -Level "DEBUG"
        return $false
    }
}

try {
    $latestVirtualBoxVersion = Get-LatestVersion -Url "https://download.virtualbox.org/virtualbox/LATEST.TXT"
    Write_LogEntry -Message "Gefundene neueste VirtualBox-Version: $($latestVirtualBoxVersion)" -Level "INFO"
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen der LATEST.TXT: $($_)" -Level "ERROR"
    $latestVirtualBoxVersion = $null
}

if ($latestVirtualBoxVersion) {
    try {
        $md5SumsContent = (Invoke-WebRequest -Uri "https://download.virtualbox.org/virtualbox/$latestVirtualBoxVersion/MD5SUMS" -UseBasicParsing).Content
        Write_LogEntry -Message "MD5SUMS Inhalt abgerufen für Version: $($latestVirtualBoxVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen der MD5SUMS für $($latestVirtualBoxVersion): $($_)" -Level "ERROR"
        $md5SumsContent = ""
    }
} else {
    $md5SumsContent = ""
    Write_LogEntry -Message "Keine Online-Version ermittelt, md5Sums übersprungen." -Level "WARNING"
}

$virtualBoxInstallerFilename = [regex]::Match($md5SumsContent, 'VirtualBox-.*-Win\.exe').Value
#$extensionPackFilename = [regex]::Match($md5SumsContent, 'Oracle_VM_VirtualBox_Extension_Pack-.*\.vbox-extpack').Value
$extensionPackFilename = [regex]::Match($md5SumsContent, 'Oracle_VirtualBox_Extension_Pack-\d+\.\d+\.\d+-\d+\.vbox-extpack').Value
Write_LogEntry -Message "Extrahierte Dateinamen: Installer=$($virtualBoxInstallerFilename); ExtensionPack=$($extensionPackFilename)" -Level "DEBUG"

# Use regular expression to extract the version number
$fullVersionNumber = $virtualBoxInstallerFilename -match '\d+\.\d+\.\d+-\d+'

# Check if a match was found and get the matched value
if ($fullVersionNumber) {
    $fullVersionNumber = $Matches[0] -replace '-', '.'
    Write_LogEntry -Message "Vollständige Version aus Dateiname extrahiert: $($fullVersionNumber)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Konnte vollständige Version aus Dateiname nicht extrahieren." -Level "WARNING"
}

$virtualBoxInstaller = Get-ChildItem -Path $InstallationFileVirtualBox | Select-Object -Last 1
if ($virtualBoxInstaller) {
    $existingVersion = (Get-Item $virtualBoxInstaller.FullName).VersionInfo.FileVersion
    Write_LogEntry -Message "Gefundener lokaler VirtualBox-Installer: $($virtualBoxInstaller.FullName) Version: $($existingVersion)" -Level "DEBUG"
	
	Write-Host ""
	Write-Host "Lokale Version: $existingVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $fullVersionNumber" -foregroundcolor "Cyan"
	Write-Host ""
	
	$result = Compare-VersionNumbers -existingVersion $existingVersion -fullVersionNumber $fullVersionNumber

    if ($result) {
        #Write-Host "Newer version of VirtualBox installer available. Downloading..."
		$downloadPath = "$InstallationFolder\$virtualBoxInstallerFilename"
        DownloadFile -Url "https://download.virtualbox.org/virtualbox/$latestVirtualBoxVersion/$virtualBoxInstallerFilename" -Path $downloadPath
		
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
			try {
				Remove-Item -Path $virtualBoxInstaller.FullName -Force
				Write_LogEntry -Message "Alte VirtualBox-Installer-Datei entfernt: $($virtualBoxInstaller.FullName)" -Level "DEBUG"
			} catch {
				Write_LogEntry -Message "Fehler beim Entfernen der alten Installer-Datei: $($_)" -Level "WARNING"
			}

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
			Write_LogEntry -Message "$($ProgramName) Installer aktualisiert: $($downloadPath)" -Level "SUCCESS"
		} else {
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
			Write_LogEntry -Message "Download fehlgeschlagen für $($virtualBoxInstallerFilename)" -Level "ERROR"
		}		
    } else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Update für $($ProgramName) verfügbar (Local: $($existingVersion), Remote: $($fullVersionNumber))" -Level "INFO"
    }
} else {
	$downloadPath = "$InstallationFolder\$virtualBoxInstallerFilename"
    Write-Host "$ProgramName wurde nicht gefunden. Runterladen..." -foregroundcolor "DarkGray"
    Write_LogEntry -Message "$($ProgramName) Installer nicht gefunden lokal; starte Download: $($downloadPath)" -Level "INFO"
    DownloadFile -Url "https://download.virtualbox.org/virtualbox/$latestVirtualBoxVersion/$virtualBoxInstallerFilename" -Path $downloadPath

	# Check if the file was completely downloaded
	if (Test-Path $downloadPath) {
		Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
        Write_LogEntry -Message "$($ProgramName) Installer heruntergeladen: $($downloadPath)" -Level "SUCCESS"
	} else {
		Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
        Write_LogEntry -Message "Download des neuen Installers fehlgeschlagen: $($downloadPath)" -Level "ERROR"
	}
}

$extensionPack = Get-ChildItem -Path $InstallationFileVirtualBoxExtension | Select-Object -Last 1
if ($extensionPack) {
    $extensionPackVersion = $extensionPack.Name -replace 'Oracle_VirtualBox_Extension_Pack-(\d+\.\d+\.\d+-\d+)\.vbox-extpack', '$1'
    $existingVersion = $extensionPackVersion -replace '-', '.'
    Write_LogEntry -Message "Gefundene lokale Extension Pack Datei: $($extensionPack.FullName); Version: $($existingVersion)" -Level "DEBUG"
	
	Write-Host ""
	Write-Host "Lokale Version: $existingVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $fullVersionNumber" -foregroundcolor "Cyan"
	Write-Host ""
	
	$result = Compare-VersionNumbers -existingVersion $existingVersion -fullVersionNumber $fullVersionNumber
	
    #$existingVersion = (Get-Item $extensionPack.FullName).VersionInfo.FileVersion
    if ($result) {
        #Write-Host "Newer version of VirtualBox Extension Pack available. Downloading..."
		$downloadPath = "$InstallationFolder\$extensionPackFilename"
        DownloadFile -Url "https://download.virtualbox.org/virtualbox/$latestVirtualBoxVersion/$extensionPackFilename" -Path $downloadPath
		
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
			try {
				Remove-Item -Path $extensionPack.FullName -Force
				Write_LogEntry -Message "Altes Extension Pack entfernt: $($extensionPack.FullName)" -Level "DEBUG"
			} catch {
				Write_LogEntry -Message "Fehler beim Entfernen des alten Extension Pack: $($_)" -Level "WARNING"
			}

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "Extension Pack heruntergeladen: $($downloadPath)" -Level "SUCCESS"
		} else {
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
            Write_LogEntry -Message "Download des Extension Pack fehlgeschlagen: $($downloadPath)" -Level "ERROR"
		}
    } else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Update für Extension Pack notwendig (Local: $($existingVersion), Remote: $($fullVersionNumber))" -Level "INFO"
    }
} else {
	$downloadPath = "$InstallationFolder\$extensionPackFilename"
    Write-Host "$ProgramName wurde nicht gefunden. Runterladen..." -foregroundcolor "DarkGray"
    Write_LogEntry -Message "Kein lokales Extension Pack gefunden; starte Download: $($downloadPath)" -Level "INFO"
    DownloadFile -Url "https://download.virtualbox.org/virtualbox/$latestVirtualBoxVersion/$extensionPackFilename" -Path $downloadPath
	
	# Check if the file was completely downloaded
	if (Test-Path $downloadPath) {
		Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
        Write_LogEntry -Message "Extension Pack heruntergeladen: $($downloadPath)" -Level "SUCCESS"
	} else {
		Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
        Write_LogEntry -Message "Download Extension Pack fehlgeschlagen: $($downloadPath)" -Level "ERROR"
	}
}

Write-Host ""

#Check Installed Version / Install if neded
$ProgramName = "Oracle VM VirtualBox"
Write_LogEntry -Message "Setze ProgramName für Installationsprüfung: $($ProgramName)" -Level "DEBUG"
$virtualBoxInstaller = Get-ChildItem -Path $InstallationFileVirtualBox | Select-Object -Last 1
if ($virtualBoxInstaller) {
	try {
		$localVersion = (Get-Item $virtualBoxInstaller.FullName).VersionInfo.FileVersion
		Write_LogEntry -Message "Lokale Installationsdatei Version ermittelt: $($localVersion)" -Level "DEBUG"
	} catch {
		Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Installationsdatei Version: $($_)" -Level "WARNING"
		$localVersion = $null
	}
} else {
	Write_LogEntry -Message "Keine lokale Installationsdatei für VirtualBox gefunden mit Muster: $($InstallationFileVirtualBox)" -Level "DEBUG"
	$localVersion = $null
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

#$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

#$Path = foreach ($RegPath in $RegistryPaths) {
    #if (Test-Path $RegPath) {
        #Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    #}
#}

$VirtualBoxPath = "C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"
Write_LogEntry -Message "Prüfe Existenz VirtualBox.exe unter: $($VirtualBoxPath)" -Level "DEBUG"

# Check if the file exists
if (Test-Path $VirtualBoxPath) {
    # Get the product version from the file properties
    $versionInfo = (Get-Item $VirtualBoxPath).VersionInfo
    $installedVersion = $versionInfo.ProductVersion
    Write_LogEntry -Message "Gefundene installierte VirtualBox.exe Version: $($installedVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "VirtualBox.exe nicht gefunden: $($VirtualBoxPath)" -Level "INFO"
    $installedVersion = $null
}

#if ($Path -ne $null) {
if ($null -ne $installedVersion) {
    #$DisplayVersion = $Path.DisplayVersion | Select-Object -First 1
    #$VersionRevision = $Path.VersionRevision | Select-Object -First 1
	#$installedVersion = $DisplayVersion + "." + $VersionRevision
	
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	Write_LogEntry -Message "$($ProgramName) installiert; InstalledVersion=$($installedVersion); LocalInstallerVersion=$($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$InstallVirtBox = $true
        Write_LogEntry -Message "InstallVirtBox = $($InstallVirtBox) (Update erforderlich)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$InstallVirtBox = $false
        Write_LogEntry -Message "InstallVirtBox = $($InstallVirtBox) (Version aktuell)" -Level "DEBUG"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$InstallVirtBox = $false
        Write_LogEntry -Message "InstallVirtBox = $($InstallVirtBox) (installierte Version neuer)" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$InstallVirtBox = $false
    Write_LogEntry -Message "$($ProgramName) nicht installiert oder Version nicht ermittelbar; InstallVirtBox = $($InstallVirtBox)" -Level "INFO"
}
Write-Host ""

$ProgramName = "VirtualBox Extension Pack"
Write_LogEntry -Message "Setze ProgramName für Extension Pack Prüfung: $($ProgramName)" -Level "DEBUG"
$InstalledFile = ""
$extensionPack = Get-ChildItem -Path $InstallationFileVirtualBoxExtension | Select-Object -Last 1
if ($extensionPack) {
    $extensionPackVersion = $extensionPack.Name -replace 'Oracle_VirtualBox_Extension_Pack-(\d+\.\d+\.\d+-\d+)\.vbox-extpack', '$1'
    $localVersion = $extensionPackVersion -replace '-', '.'
    Write_LogEntry -Message "Gefundene lokale Extension Pack Datei: $($extensionPack.FullName); Version: $($localVersion)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein lokales Extension Pack gefunden mit Muster: $($InstallationFileVirtualBoxExtension)" -Level "DEBUG"
    $localVersion = $null
}

$ProgramFiles = [Environment]::GetFolderPath('ProgramFiles')
$ExtensionPackFolderPath = Join-Path -Path $ProgramFiles -ChildPath "Oracle\VirtualBox\ExtensionPacks"
Write_LogEntry -Message "Ermittle ExtensionPack Pfad: $($ExtensionPackFolderPath)" -Level "DEBUG"

# Check if the ExtensionPacks folder exists
if (Test-Path -Path $ExtensionPackFolderPath) {
    # Get the folder that matches *VirtualBox_Extension_Pack
	#$InstalledFile = Join-Path -Path $ExtensionPackFolderPath -ChildPath "\Oracle_VM_VirtualBox_Extension_Pack\ExtPack.xml"
    $extensionPackFolder = Get-ChildItem -Path $ExtensionPackFolderPath -Directory | Where-Object { $_.Name -like '*VirtualBox_Extension_Pack*' }

    if ($extensionPackFolder) {
        # If the folder exists, get the full path to ExtPack.xml
        $InstalledFile = Join-Path -Path $extensionPackFolder.FullName -ChildPath "ExtPack.xml"
        Write_LogEntry -Message "Gefundenes InstalledFile ExtPack.xml: $($InstalledFile)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Kein ExtensionPack-Ordner unter $($ExtensionPackFolderPath) gefunden." -Level "DEBUG"
    }
} else {
    Write_LogEntry -Message "ExtensionPackFolderPath existiert nicht: $($ExtensionPackFolderPath)" -Level "DEBUG"
}

if (-not [string]::IsNullOrEmpty($InstalledFile)) {
    if (Test-Path $InstalledFile) {
		try {
			# Load the XML content from the file
			$xmlContent = [xml](Get-Content $InstalledFile)

			# Extract the version and revision
			$version = $xmlContent.VirtualBoxExtensionPack.Version
			$revision = $version.revision

			# Combine version and revision
			$installedVersion = "$($version.'#text').$revision"
			Write_LogEntry -Message "Ermittelte Extension Pack installierte Version aus ExtPack.xml: $($installedVersion)" -Level "DEBUG"
		} catch {
			Write_LogEntry -Message "Fehler beim Lesen von ExtPack.xml: $($_)" -Level "ERROR"
			$installedVersion = $null
		}
	} else {
		Write_LogEntry -Message "InstalledFile existiert nicht: $($InstalledFile)" -Level "DEBUG"
		$installedVersion = $null
	}
} else {
	Write_LogEntry -Message "Kein InstalledFile definiert; überspringe Extension Pack Installed-Versionsermittlung." -Level "DEBUG"
	$installedVersion = $null
}

if ($null -ne $installedVersion) {	
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	Write_LogEntry -Message "$($ProgramName) installiert; InstalledVersion=$($installedVersion); LocalInstallerVersion=$($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$InstallVirtBoxExPack = $true
        Write_LogEntry -Message "InstallVirtBoxExPack = $($InstallVirtBoxExPack) (Update erforderlich)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$InstallVirtBoxExPack = $false
        Write_LogEntry -Message "InstallVirtBoxExPack = $($InstallVirtBoxExPack) (Version aktuell)" -Level "DEBUG"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$InstallVirtBoxExPack = $false
        Write_LogEntry -Message "InstallVirtBoxExPack = $($InstallVirtBoxExPack) (installierte Version neuer)" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$InstallVirtBoxExPack = $false
    Write_LogEntry -Message "Extension Pack nicht installiert oder Version nicht ermittelbar; InstallVirtBoxExPack = $($InstallVirtBoxExPack)" -Level "INFO"
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "Starte externes Installationsscript (InstallationFlag) via $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\VirtualBoxInstall.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Rückkehr von VirtualBoxInstall.ps1 nach InstallationFlag-Aufruf" -Level "DEBUG"
}

if($InstallVirtBox -eq $true){
    Write_LogEntry -Message "Starte externes Installationsscript für VirtualBox: $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\VirtualBoxInstall.ps1" `
		-InstallVirtBox
    Write_LogEntry -Message "Rückkehr von VirtualBoxInstall.ps1 nach InstallVirtBox-Aufruf" -Level "DEBUG"
}

if($InstallVirtBoxExPack -eq $true){
    Write_LogEntry -Message "Starte externes Installationsscript für Extension Pack: $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\VirtualBoxInstall.ps1" `
		-InstallVirtBoxExPack
    Write_LogEntry -Message "Rückkehr von VirtualBoxInstall.ps1 nach InstallVirtBoxExPack-Aufruf" -Level "DEBUG"
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
