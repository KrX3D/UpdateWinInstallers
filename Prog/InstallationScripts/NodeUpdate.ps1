param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Node.js"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

# Check for existing Node.js installer in installation folder
$InstallationFilePattern = "$InstallationFolder\node-v*-x64.msi"
Write_LogEntry -Message "Suchmuster für lokale Installer: $($InstallationFilePattern)" -Level "DEBUG"
$FoundFile = Get-ChildItem $InstallationFilePattern -ErrorAction SilentlyContinue
Write_LogEntry -Message "Lokale Installer-Datei gefunden: $([bool]$FoundFile)" -Level "DEBUG"

# Initialize variables
$localVersion = "0.0.0"
$localInstaller = ""

# If a local installer exists, get its version
if ($FoundFile) {
    $InstallationFileName = $FoundFile.Name
    $localInstaller = "$InstallationFolder\$InstallationFileName"
    Write_LogEntry -Message "Lokaler Installer Pfad: $($localInstaller)" -Level "DEBUG"
    
    # Extract version from filename (e.g., node-v18.16.0-x64.msi -> 18.16.0)
    $versionMatch = [regex]::Match($InstallationFileName, 'node-v([\d\.]+)-x64\.msi')
    if ($versionMatch.Success) {
        $localVersion = $versionMatch.Groups[1].Value
        Write_LogEntry -Message "Lokale Version aus Dateiname extrahiert: $($localVersion)" -Level "INFO"
    } else {
        Write_LogEntry -Message "Version konnte nicht aus Dateiname extrahiert werden: $($InstallationFileName)" -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Kein lokaler Installer vorhanden; setze lokale Version auf Default $($localVersion)" -Level "DEBUG"
}

# Define the URL for the latest Node.js version
$nodeLatestUrl = "https://nodejs.org/download/release/latest/"
Write_LogEntry -Message "Node.js Versions-URL gesetzt: $($nodeLatestUrl)" -Level "DEBUG"

# Retrieve the download page content
try {
    Write_LogEntry -Message "Rufe Node.js Download-Seite ab: $($nodeLatestUrl)" -Level "INFO"
    $webRequest = Invoke-WebRequest -Uri $nodeLatestUrl -UseBasicParsing
    $pageContent = $webRequest.Content
    Write_LogEntry -Message "Node.js Download-Seite abgerufen (Inhalt vorhanden: $($([bool]$pageContent)))" -Level "DEBUG"

    # Extract the Windows x64 MSI installer link and version
    $pattern = 'href="[^"]*?(node-v([\d\.]+)-x64\.msi)"'
    $match = [regex]::Match($pageContent, $pattern)

    if ($match.Success) {
        $latestFileName = $match.Groups[1].Value
        $latestVersion = $match.Groups[2].Value
        $downloadUrl = "$nodeLatestUrl$latestFileName"
        Write_LogEntry -Message "Gefundene Online-Datei: $($latestFileName); Online-Version: $($latestVersion); DownloadURL: $($downloadUrl)" -Level "INFO"
        
		Write-Host ""
		Write-Host "Lokale Version: $localVersion" -ForegroundColor "Cyan"
        Write-Host "Online Version: $latestVersion" -ForegroundColor "Cyan"
        Write-Host ""

        # Compare versions
        if ([version]$latestVersion -gt [version]$localVersion) {
            Write_LogEntry -Message "Online-Version $($latestVersion) ist neuer als lokal $($localVersion); starte Download" -Level "INFO"
            Write-Host "Eine neuere Version von $ProgramName ist verfügbar. Update wird heruntergeladen..." -ForegroundColor "Yellow"
			Write-Host ""
            
            # Download the new version
            $tempFilePath = "$InstallationFolder\$latestFileName"
            Write_LogEntry -Message "Download-Zielpfad: $($tempFilePath)" -Level "DEBUG"
            
            $webClient = New-Object System.Net.WebClient
            try {
                Write_LogEntry -Message "Starte Download $($downloadUrl) nach $($tempFilePath)" -Level "INFO"
                [void](Invoke-DownloadFile -Url $downloadUrl -OutFile $tempFilePath)
                Write_LogEntry -Message "Download abgeschlossen: $($tempFilePath)" -Level "DEBUG"
            }
            catch {
                Write_LogEntry -Message "Fehler beim Herunterladen der Datei: $($($_.Exception.Message))" -Level "ERROR"
                throw
            }
            finally {
                $webClient.Dispose()
            }
            
            # Check if the file was completely downloaded
            if (Test-Path $tempFilePath) {
                # Remove the old installer if it exists
                if ($FoundFile) {
                    try {
                        Remove-Item $localInstaller -Force
                        Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localInstaller)" -Level "DEBUG"
                    } catch {
                        Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei: $($($_.Exception.Message))" -Level "WARNING"
                    }
                }
                
                Write_LogEntry -Message "$($ProgramName) wurde aktualisiert auf $($latestVersion)" -Level "SUCCESS"
                Write-Host "$ProgramName wurde aktualisiert." -ForegroundColor "Green"
				Write-Host ""
                $localInstaller = $tempFilePath
                $localVersion = $latestVersion
            } else {
                Write_LogEntry -Message "Download ist fehlgeschlagen; Datei nicht gefunden: $($tempFilePath)" -Level "ERROR"
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "Red"
            }
        } else {
            Write_LogEntry -Message "Kein Online Update verfügbar. Online: $($latestVersion); Lokal: $($localVersion)" -Level "INFO"
            Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
        }
    } else {
        Write_LogEntry -Message "Konnte die neueste Version von $($ProgramName) nicht ermitteln (Pattern nicht gefunden)." -Level "ERROR"
        Write-Host "Konnte die neueste Version von $ProgramName nicht ermitteln." -ForegroundColor "Red"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen der $($ProgramName)-Versionsinformationen: $($($_.Exception.Message))" -Level "ERROR"
    Write-Host "Fehler beim Abrufen der $ProgramName-Versionsinformationen: $($_)" -ForegroundColor "Red"
}

# Check installed version
$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Prüfung: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}
Write_LogEntry -Message "Registry-Abfrage Ergebnis vorhanden: $($([bool]$Path))" -Level "DEBUG"

$Install = $false

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write_LogEntry -Message "Gefundene installierte Version: $($installedVersion); Installationsdatei Version: $($localVersion)" -Level "INFO"
    Write-Host "$ProgramName ist installiert." -ForegroundColor "Green"
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor "Cyan"
    
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write_LogEntry -Message "Installierte Version $($installedVersion) ist älter als lokale Datei $($localVersion); Install = True" -Level "INFO"
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "Magenta"
        $Install = $true
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write_LogEntry -Message "Installierte Version entspricht lokaler Datei; Install = False" -Level "DEBUG"
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor "DarkGray"
        $Install = $false
    } else {
        Write_LogEntry -Message "Installierte Version ist neuer als lokale Datei; Install = False" -Level "WARNING"
        $Install = $false
    }
} else {
    #Write-Host "$ProgramName ist nicht installiert." -ForegroundColor "Yellow"
    Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden; Install = False" -Level "INFO"
    $Install = $false
}
Write_LogEntry -Message "Install-Flag nach Prüfung: $($Install)" -Level "DEBUG"
Write-Host ""

# Install if needed
if ($InstallationFlag) {
    Write_LogEntry -Message "InstallationFlag gesetzt; rufe Installationsskript mit Flag auf: $($Serverip)\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Installationsskript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1" -Level "DEBUG"
} elseif ($Install -eq $true) {
    Write_LogEntry -Message "Starte externes Installationsskript (Update): $($Serverip)\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1"
    Write_LogEntry -Message "Externes Installationsskript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\NodeJSInstallation.ps1" -Level "DEBUG"
}
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"
Write-Host ""

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
