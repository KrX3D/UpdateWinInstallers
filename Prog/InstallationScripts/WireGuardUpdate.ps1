param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WireGuard"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$NetworkShareDaten  = $config.NetworkShareDaten

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
    Write-Host "Lokale Version: $localVersion" -ForegroundColor "Cyan"
    Write-Host "Online Version: $latestVersion" -ForegroundColor "Cyan"
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

            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor "Green"
            Write_LogEntry -Message "$($ProgramName) Update erfolgreich: $($downloadPath)" -Level "SUCCESS"
        } else {
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "Red"
            Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath) nicht gefunden nach Download" -Level "ERROR"
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
        Write_LogEntry -Message "Kein Update verfügbar: Online $($latestVersion) <= Local $($localVersion)" -Level "INFO"
    }
} else {
    Write_LogEntry -Message "Keine WireGuard-Installationsdatei gefunden im Pfad: $($installerPath)" -Level "WARNING"
}

Write-Host ""

# Check Installed Version / Install if needed
$installerFile = Get-InstallerFilePath -PathPattern $installerPath
if ($installerFile) {
    $versionPattern = 'wireguard-amd64-(\d+\.\d+\.\d+)\.msi'
    $localVersion = Get-InstallerFileVersion -FilePath $installerFile.FullName -FileNameRegex $versionPattern -Source FileName
    Write_LogEntry -Message "Ermittelte lokale Dateiversion nach erneutem Scan: $($localVersion) (Datei: $($installerFile.FullName))" -Level "DEBUG"
} else {
    Write_LogEntry -Message "Keine Installationsdatei gefunden beim zweiten Scan: $($installerPath)" -Level "DEBUG"
}

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
    Write-Host "$ProgramName ist installiert." -ForegroundColor "Green"
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor "Cyan"
    Write_LogEntry -Message "Gefundene installierte Version aus Registry: $($installedVersion); Datei-Version: $($localVersion)" -Level "INFO"

    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "Magenta"
        $Install = $true
        Write_LogEntry -Message "Install erforderlich: Registry $($installedVersion) < Datei $($localVersion)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor "DarkGray"
        $Install = $false
        Write_LogEntry -Message "Install nicht erforderlich: InstalledVersion == LocalVersion ($($localVersion))" -Level "INFO"
    } else {
        $Install = $false
        Write_LogEntry -Message "InstalledVersion ($($installedVersion)) > LocalVersion ($($localVersion)); keine Aktion" -Level "WARNING"
    }
} else {
    $Install = $false
    Write_LogEntry -Message "Keine Registry-Einträge für $($ProgramName) gefunden" -Level "DEBUG"
}
Write-Host ""

# Install if needed
if ($InstallationFlag) {
    Write_LogEntry -Message "Starte externes Installationsskript aufgrund InstallationFlag" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$NetworkShareDaten\Prog\InstallationScripts\Installation\WireGuardInstall.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WireGuardInstall.ps1 mit -InstallationFlag" -Level "DEBUG"
} elseif ($Install -eq $true) {
    Write_LogEntry -Message "Starte externes Installationsskript aufgrund Install=true" -Level "INFO"
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$NetworkShareDaten\Prog\InstallationScripts\Installation\WireGuardInstall.ps1"
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WireGuardInstall.ps1" -Level "DEBUG"
}
Write-Host ""

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
