param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Autodesk Fusion"
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

$InstallationFolder = "$InstallationFolder\3D"
$localInstaller = Join-Path $InstallationFolder "Fusion_Client_Downloader.exe"
$TempDownloadPath = "$env:TEMP\Fusion_Client_Downloader.exe"
$DownloadUrl = "https://dl.appstreaming.autodesk.com/production/installers/Fusion%20Client%20Downloader.exe"

Write_LogEntry -Message "InstallationFolder: $($InstallationFolder); LocalInstaller: $($localInstaller); TempDownloadPath: $($TempDownloadPath)" -Level "DEBUG"
Write_LogEntry -Message "DownloadUrl: $($DownloadUrl)" -Level "DEBUG"

function Get-LatestOnlineVersion {
    try {
        Write_LogEntry -Message "Get-LatestOnlineVersion: Starte Web-Request an http://autode.sk/whatsnew" -Level "INFO"
        $webRequest = Invoke-WebRequest -Uri "http://autode.sk/whatsnew" -UseBasicParsing
        $versionMatches = [regex]::Matches($webRequest.Content, "v\.(\d+\.\d+\.\d+)")
        if ($versionMatches.Count -eq 0) {
			Write_LogEntry -Message "Get-LatestOnlineVersion: Keine gültige Version auf der Webseite gefunden." -Level "WARNING"
			Write-Host "Keine gültige Version gefunden." -ForegroundColor Red
            return $null
        }
        # Get the highest version
        $latestVersion = $versionMatches | ForEach-Object { $_.Groups[1].Value } |
            Sort-Object { [version]$_ } -Descending |
            Select-Object -First 1
        Write_LogEntry -Message "Get-LatestOnlineVersion: Gefundene Online-Version: $($latestVersion)" -Level "INFO"
        return $latestVersion
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen der Online-Version: $($_)" -Level "ERROR"
        Write-Host "Fehler beim Abrufen der Online-Version: $_" -ForegroundColor Red
        exit
    }
}

function Get-LocalVersion {
    if (-Not (Test-Path $localInstaller)) {
        Write_LogEntry -Message "Get-LocalVersion: Lokale Datei nicht gefunden: $($localInstaller)" -Level "WARNING"
        #Write-Host "Lokale Datei nicht gefunden: $localInstaller" -ForegroundColor Yellow
        return $null
    }
    $localVersion = (Get-Item $localInstaller).VersionInfo.FileVersion
    Write_LogEntry -Message "Get-LocalVersion: Gefundene lokale FileVersion: $($localVersion)" -Level "DEBUG"
    $parts = $localVersion.Split('.')
    if ($parts.Length -lt 3) {
        Write_LogEntry -Message "Get-LocalVersion: Ungültige lokale Version: $($localVersion)" -Level "ERROR"
        Write-Host "Ungültige lokale Version." -ForegroundColor Red
        return $null
    }
    # Reorder to match online structure: 2602.0.71
    $reordered = "$($parts[1]).$($parts[2]).$($parts[0])"
    Write_LogEntry -Message "Get-LocalVersion: Umgeordnete lokale Version: $($reordered)" -Level "DEBUG"
    return $reordered
}

function Compare-Versions($local, $online) {
    Write_LogEntry -Message "Compare-Versions: Vergleiche Local: $($local) mit Online: $($online)" -Level "DEBUG"
    $v1 = [version]$local
    $v2 = [version]$online
    $result = ($v2 -gt $v1)
    Write_LogEntry -Message "Compare-Versions: Ergebnis (online > local) = $($result)" -Level "DEBUG"
    return $result
}

function Download-And-Replace {
    try {
        Write_LogEntry -Message "Download-And-Replace: Starte Download von $($DownloadUrl) nach $($TempDownloadPath)" -Level "INFO"
        [void](Invoke-DownloadFile -Url $DownloadUrl -OutFile $TempDownloadPath)
        Write_LogEntry -Message "Download-And-Replace: Download beendet. Prüfe Existenz: $($TempDownloadPath)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Download-And-Replace: Download fehlgeschlagen: $($_)" -Level "ERROR"
        Write-Host "Download fehlgeschlagen: $_" -ForegroundColor Red
        return
    }

    if (-Not (Test-Path $TempDownloadPath)) {
		Write_LogEntry -Message "Download-And-Replace: Temp-Download-Datei nicht gefunden: $($TempDownloadPath)" -Level "ERROR"
		Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
        return
    }

    try {
        if (Test-Path $localInstaller) {
            Write_LogEntry -Message "Download-And-Replace: Lokale Installer-Datei existiert: $($localInstaller). Entferne alte Datei." -Level "INFO"
            Remove-Item $localInstaller -Force
			Move-Item -Path (Join-Path $env:TEMP "Fusion_Client_Downloader.exe") -Destination $InstallationFolder -Force
			Write_LogEntry -Message "Download-And-Replace: Neue Datei verschoben nach $($InstallationFolder)" -Level "INFO"
			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
		} else {
            Write_LogEntry -Message "Download-And-Replace: Lokale Installer-Datei nicht vorhanden zum Ersetzen: $($localInstaller)" -Level "ERROR"
			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
		}
    } catch {
        Write_LogEntry -Message "Download-And-Replace: Fehler beim Ersetzen: $($_)" -Level "ERROR"
        Write-Host "Fehler beim Ersetzen: $_" -ForegroundColor Red
    }
}

# MAIN EXECUTION
$localVersion = Get-LocalVersion
$onlineVersion = Get-LatestOnlineVersion

Write_LogEntry -Message "Main: Lokale Version: $($localVersion); Online Version: $($onlineVersion)" -Level "INFO"

Write-Host ""
Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
Write-Host ""

if ($localVersion -eq $null -or (Compare-Versions -local $localVersion -online $onlineVersion)) {
    Write_LogEntry -Message "Main: Update erforderlich oder lokale Version unbekannt. Local: $($localVersion); Online: $($onlineVersion)" -Level "INFO"
    Write-Host "Update erforderlich!" -ForegroundColor Yellow
    Write_LogEntry -Message "Main: Rufe Download-And-Replace auf." -Level "DEBUG"
    Download-And-Replace
    Write_LogEntry -Message "Main: Download-And-Replace zurück." -Level "DEBUG"
} else {
	Write_LogEntry -Message "Main: Kein Online Update verfügbar. $($ProgramName) ist aktuell." -Level "INFO"
	Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
}

#Check Installed Version / Install if neded
$localVersion = Get-LocalVersion
Write_LogEntry -Message "Check-Installed: Erneute lokale Version: $($localVersion)" -Level "DEBUG"

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Prüfung: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
	Write_LogEntry -Message "Check-Installed: Gefundene installierte Version von $($ProgramName): $($installedVersion)" -Level "INFO"

	Write-Host ""
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
        Write_LogEntry -Message "Check-Installed: Installierte Version ($($installedVersion)) ist älter als lokale Datei ($($localVersion)). Markiere Install=true" -Level "INFO"
		$Install = $true
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Check-Installed: Installierte Version ist aktuell: $($installedVersion)" -Level "DEBUG"
		$Install = $false
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
        Write_LogEntry -Message "Check-Installed: Installierte Version ($($installedVersion)) ist neuer als lokale Datei ($($localVersion)). Kein Update nötig." -Level "WARNING"
		$Install = $false
    }
} else {
    Write_LogEntry -Message "Check-Installed: $($ProgramName) nicht in der Registry gefunden. Setze Install-Flag auf $($false)." -Level "INFO"
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
}
Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "Install-Schritt: InstallationFlag gesetzt. Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Fusion360Installation.ps1 mit -InstallationFlag" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\Fusion360Installation.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Install-Schritt: Externes Installations-Skript mit -InstallationFlag aufgerufen." -Level "DEBUG"
}
elseif($Install -eq $true){
    Write_LogEntry -Message "Install-Schritt: Install-Flag true. Starte externes Installations-Skript: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Fusion360Installation.ps1" -Level "INFO"
	#Uninstall + install
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\Fusion360Installation.ps1"
    Write_LogEntry -Message "Install-Schritt: Externes Installations-Skript ohne Parameter aufgerufen." -Level "DEBUG"
}
Write-Host ""
Write_LogEntry -Message "Skript-Ende erreicht. Vor Footer." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
