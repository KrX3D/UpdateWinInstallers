param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Agent Ransack"
$ScriptType = "Update"

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

Write_LogEntry -Message "Prüfe Vorhandensein der Konfigurationsdatei: $($configPath)" -Level "DEBUG"

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

Write_LogEntry -Message "ProgramName gesetzt auf: $($ProgramName)" -Level "DEBUG"

$localFileWildcard = "agentransack*.msi"
Write_LogEntry -Message "Lokales Datei-Wildcard gesetzt auf: $($localFileWildcard)" -Level "DEBUG"

$TempFolder = "$env:TEMP"
Write_LogEntry -Message "Temp-Ordner: $($TempFolder)" -Level "DEBUG"

$onlineVersionUrl = "https://www.mythicsoft.com/agentransack/"
Write_LogEntry -Message "Online-Version URL gesetzt auf: $($onlineVersionUrl)" -Level "DEBUG"

# Get the latest version from the online website
Write_LogEntry -Message "Rufe Online-Seite $($onlineVersionUrl) ab, um Version zu ermitteln." -Level "INFO"
$onlineVersionHtml = Invoke-WebRequest -Uri $onlineVersionUrl -UseBasicParsing
if ($null -ne $onlineVersionHtml) {
    Write_LogEntry -Message "Online-Seite abgerufen. Content-Länge: $($onlineVersionHtml.Content.Length)" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Fehler: Keine Antwort von $($onlineVersionUrl)" -Level "WARNING"
}

$onlineVersion = $null
try {
    $onlineVersion = $onlineVersionHtml.Content | Select-String -Pattern "agentransack_(\d+)" | ForEach-Object { $_.Matches.Groups[1].Value } | Select-Object -First 1
    Write_LogEntry -Message "Online-Version extrahiert: $($onlineVersion)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Extrahieren der Online-Version: $($_)" -Level "ERROR"
}

# Get the local file matching the wildcard pattern
Write_LogEntry -Message "Suche lokale Datei im Installationsordner mit Filter: $($localFileWildcard)" -Level "DEBUG"
$localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard | Select-Object -First 1 -ExpandProperty FullName

if (-not $localFilePath) {
    Write_LogEntry -Message "Keine lokale Datei gefunden für Pattern: $($localFileWildcard) in $($InstallationFolder)" -Level "WARNING"
    #Write-Host "No local file found matching the wildcard pattern: $localFileWildcard"
} else {
    Write_LogEntry -Message "Lokale Datei gefunden: $($localFilePath)" -Level "INFO"

    # Retrieve local file version from the filename
    try {
        $localVersion = [System.IO.Path]::GetFileNameWithoutExtension($localFilePath) -replace 'agentransack_', '' -replace 'x64_', ''
        Write_LogEntry -Message "Lokale Version aus Dateiname extrahiert: $($localVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Version aus Dateiname $($localFilePath): $($_)" -Level "ERROR"
    }

    # Compare versions
    if ($null -ne $onlineVersion -and $null -ne $localVersion) {
        try {
            $isLocalVersionNewer = [int]$onlineVersion -gt [int]$localVersion
            Write_LogEntry -Message "Versionsvergleich: Online $($onlineVersion) vs Lokal $($localVersion) => isLocalVersionNewer = $($isLocalVersionNewer)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Vergleich der Versionen Online:$($onlineVersion) Lokal:$($localVersion): $($_)" -Level "ERROR"
            $isLocalVersionNewer = $false
        }
    } else {
        Write_LogEntry -Message "Versionen konnten nicht verglichen werden (Online:$($onlineVersion) Lokal:$($localVersion))." -Level "WARNING"
        $isLocalVersionNewer = $false
    }

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
    Write-Host ""

    if (!$isLocalVersionNewer) {
        Write_LogEntry -Message "Kein Online Update verfügbar oder lokale Version ist aktuell. Online:$($onlineVersion) Lokal:$($localVersion)" -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    } else {
        Write_LogEntry -Message "Update verfügbar: Online $($onlineVersion) > Lokal $($localVersion). Starte Download." -Level "INFO"
        # Download the newer version
        $downloadUrl = "https://download.mythicsoft.com/flp/$onlineVersion/agentransack_x64_msi_$onlineVersion.zip"
        $downloadPath = Join-Path $TempFolder "agentransack_x64_msi_$onlineVersion.zip"
        Write_LogEntry -Message "Download-URL: $($downloadUrl)" -Level "DEBUG"
        Write_LogEntry -Message "Zieldatei für Download: $($downloadPath)" -Level "DEBUG"
        
        try {
            #Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($downloadUrl, $downloadPath)
            $webClient.Dispose()
            Write_LogEntry -Message "Download abgeschlossen (temporär): $($downloadPath)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Herunterladen von $($downloadUrl): $($_)" -Level "ERROR"
        }
		
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
            Write_LogEntry -Message "Downloaddatei existiert: $($downloadPath)" -Level "SUCCESS"
			#Write-Host "Removing the old version..."
            try {
                Remove-Item -Path $localFilePath -Force
                Write_LogEntry -Message "Alte Datei entfernt: $($localFilePath)" -Level "INFO"
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($localFilePath): $($_)" -Level "ERROR"
            }
			
			# Extract the downloaded version
            try {
                Write_LogEntry -Message "Entpacke $($downloadPath) nach $($InstallationFolder)" -Level "INFO"
                Expand-Archive -Path $downloadPath -DestinationPath $InstallationFolder -Force
                Write_LogEntry -Message "Entpacken erfolgreich: $($downloadPath) -> $($InstallationFolder)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "Fehler beim Entpacken $($downloadPath): $($_)" -Level "ERROR"
            }
			
			#Write-Host "Removing the zip file..."
            try {
                Remove-Item -Path $downloadPath -Force
                Write_LogEntry -Message "Zip-Datei entfernt: $($downloadPath)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen der Zip-Datei $($downloadPath): $($_)" -Level "WARNING"
            }

            Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) wurde aktualisiert. Neue Dateien im Ordner: $($InstallationFolder)" -Level "SUCCESS"
		} else {
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
            Write_LogEntry -Message "Download fehlgeschlagen, Datei nicht gefunden: $($downloadPath)" -Level "ERROR"
		}
		
    }
}

Write-Host ""

#Check Installed Version / Install if neded
Write_LogEntry -Message "Ermittle lokale Datei für Installations-Check mit Filter: $($localFileWildcard) in $($InstallationFolder)" -Level "DEBUG"
$localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard | Select-Object -First 1 -ExpandProperty FullName

if ($null -ne $localFilePath) {
    try {
        $localVersion = [System.IO.Path]::GetFileNameWithoutExtension($localFilePath) -replace 'agentransack_', '' -replace 'x64_', ''
        Write_LogEntry -Message "Lokale Installationsdatei: $($localFilePath), Version: $($localVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Ermitteln der lokalen Installationsversion aus $($localFilePath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Keine lokale Installationsdatei für Check gefunden mit Pattern: $($localFileWildcard)" -Level "WARNING"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Prüfung gesetzt: $($RegistryPaths -join ', ')" -Level "DEBUG"

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
	$versionParts = $installedVersion.Split('.')
	$installedVersion = $versionParts[2]

    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	Write_LogEntry -Message "Programm installiert. Installierte Version: $($installedVersion). Lokale Installationsdatei Version: $($localVersion)" -Level "INFO"
	
    if ($installedVersion -lt $localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Installationsentscheidung: Update erforderlich (Install = $($Install))." -Level "INFO"
    } elseif ($installedVersion -eq $localVersion) {
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
    $installerPath = "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1"
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte externes Installationsskript: $($installerPath) mittels $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Aufruf externes Skript beendet: $($installerPath)" -Level "DEBUG"
}
elseif($Install -eq $true){
    $installerPath2 = "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1"
    Write_LogEntry -Message "Install=true. Starte externes Installationsskript: $($installerPath2) mittels $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1"
    Write_LogEntry -Message "Aufruf externes Skript beendet: $($installerPath2)" -Level "DEBUG"
}

Write-Host ""

# ===== Logger-Footer (BEGIN) =====
Write_LogEntry -Message "Script beendet." -Level "INFO"
Finalize_LogSession
# ===== Logger-Footer (END) =====
