param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Google Chrome"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PSScriptRoot: $($PSScriptRoot)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

#https://github.com/MoeClub/Chrome

$localFilePath = "$InstallationFolder\GoogleChromeStandaloneEnterprise64_*.msi"
Write_LogEntry -Message "Suche lokale Chrome-Installer unter: $($localFilePath)" -Level "DEBUG"

# Get the local file version from the filename
$localFile = Get-ChildItem -Path $localFilePath | Select-Object -Last 1
Write_LogEntry -Message ("Gefundene lokale Datei: " + $([string]($localFile | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue))) -Level "DEBUG"
$localVersion = ($localFile.Name -split '_')[1] -replace '\.msi$'
Write_LogEntry -Message "Lokale Datei-Version extrahiert: $($localVersion)" -Level "DEBUG"

# Retrieve the latest version from the API endpoint
#$apiUrl = "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions/all/releases"
$apiUrl = "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions"
Write_LogEntry -Message "Rufe Google Version API ab: $($apiUrl)" -Level "INFO"
$apiResponse = Invoke-RestMethod -Uri $apiUrl
Write_LogEntry -Message "API-Antwort empfangen; Versionsanzahl: $($($apiResponse.versions).Count)" -Level "DEBUG"

#if ($apiResponse.releases -is [array] -and $apiResponse.releases.Length -gt 0) {
if ($apiResponse.versions -is [array] -and $apiResponse.versions.Length -gt 0) {
    # Get the latest version from the first release in the response
    #$latestVersion = $apiResponse.releases[0].version
    $latestVersion = $apiResponse.versions[0].version
    Write_LogEntry -Message "Ermittelte Online-Version: $($latestVersion)" -Level "INFO"

	Write-Host ""
	Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
	Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
	Write-Host ""
		
    # Compare the local and online file versions
    if ([version]$latestVersion -gt [version]$localVersion) {
        Write_LogEntry -Message "Update verfügbar: Online $($latestVersion) > Lokal $($localVersion)" -Level "INFO"

        # Construct the download link with the latest version
        $downloadLink = "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
        Write_LogEntry -Message "Verwende Download-Link: $($downloadLink)" -Level "DEBUG"

        # Download the updated installer
        $downloadPath = "$InstallationFolder\GoogleChromeStandaloneEnterprise64_$latestVersion.msi"
        Write_LogEntry -Message "Zielpfad für Download: $($downloadPath)" -Level "DEBUG"
		
        #Invoke-WebRequest -Uri $downloadLink -OutFile $downloadPath		
        Write_LogEntry -Message "Starte Download von $($downloadLink) nach $($downloadPath)" -Level "INFO"
        $webClient = New-Object System.Net.WebClient
        [void](Invoke-DownloadFile -Url $downloadLink -OutFile $downloadPath)
        $webClient.Dispose()
        Write_LogEntry -Message "Download abgeschlossen; prüfe Dateiexistenz: $($downloadPath)" -Level "DEBUG"
		
		# Check if the file was completely downloaded
		if (Test-Path $downloadPath) {
            Write_LogEntry -Message "Download erfolgreich: $($downloadPath). Entferne alte Datei: $($localFile.FullName)" -Level "INFO"
			# Remove the old installer
			Remove-Item -Path $localFile.FullName -Force
            Write_LogEntry -Message "Alte Datei entfernt: $($localFile.FullName)" -Level "DEBUG"

			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) Update abgeschlossen: Neue Version $($latestVersion)" -Level "SUCCESS"
		} else {
            Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath) wurde nicht gefunden." -Level "ERROR"
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
		}
    }
    else {
        Write_LogEntry -Message "Kein Online Update verfügbar. $($ProgramName) ist aktuell (Online: $($latestVersion); Lokal: $($localVersion))." -Level "INFO"
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    }
}
else {
    Write_LogEntry -Message "Konnte neueste Version nicht aus API-Antwort extrahieren." -Level "WARNING"
    #Write-Host "Failed to retrieve the latest version from the API endpoint."
}

Write-Host ""
Write_LogEntry -Message "Erneute Prüfung lokaler Dateien für Installationsstatus." -Level "DEBUG"

#Check Installed Version / Install if neded
$localFile = Get-ChildItem -Path $localFilePath | Select-Object -Last 1
Write_LogEntry -Message ("Gefundene lokale Datei für Install: " + $([string]($localFile | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue))) -Level "DEBUG"
$localVersion = ($localFile.Name -split '_')[1] -replace '\.msi$'
Write_LogEntry -Message "Lokale Datei-Version (erneut) extrahiert: $($localVersion)" -Level "DEBUG"

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Zu prüfende Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

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
    Write_LogEntry -Message "Gefundene installierte Version: $($installedVersion); Installationsdatei Version: $($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist älter als lokale Datei ($($localVersion)). Markiere Installation." -Level "INFO"
		$Install = $true
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		Write_LogEntry -Message "Installierte Version ist aktuell: $($installedVersion)" -Level "DEBUG"
		$Install = $false
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Datei ($($localVersion)). Kein Update nötig." -Level "WARNING"
		$Install = $false
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden. Install-Flag auf $($false) gesetzt." -Level "INFO"
	$Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1 mit Parameter -InstallationFlag" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Externes Installations-Skript mit -InstallationFlag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1" -Level "DEBUG"
}
elseif($Install -eq $true){
    Write_LogEntry -Message "Install Flag true: Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1"
    Write_LogEntry -Message "Externes Installations-Skript für Google Chrome aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\GoogleChromeInstall.ps1" -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
