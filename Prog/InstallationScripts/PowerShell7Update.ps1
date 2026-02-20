param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "PowerShell 7"
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

$localFilePath = "$InstallationFolder\PowerShell-*-win-x64.msi"
$repoOwner = "PowerShell"
$repoName = "PowerShell"
Write_LogEntry -Message "Lokaler Dateimuster-Pfad: $($localFilePath); Repo: $($repoOwner)/$($repoName)" -Level "DEBUG"

# Get the local file version from the filename
$localFile = Get-ChildItem -Path $localFilePath | Select-Object -Last 1
if ($localFile) {
    $localFileName = $localFile.Name
    $localVersion = ($localFileName -replace 'PowerShell-(\d+\.\d+\.\d+).*', '$1')
    Write_LogEntry -Message "Lokale Datei gefunden: $($localFile.FullName); Lokale Version: $($localVersion)" -Level "DEBUG"
} else {
    $localFileName = $null
    $localVersion = "0.0.0"
    Write_LogEntry -Message "Keine lokale Datei gefunden mit Muster: $($localFilePath); Setze lokale Version auf $($localVersion)" -Level "WARNING"
}

# Retrieve the latest release information from GitHub API
$apiUrl = "https://api.github.com/repos/$repoOwner/$repoName/releases/latest"
Write_LogEntry -Message "Rufe GitHub API ab: $($apiUrl)" -Level "DEBUG"

# Prepare headers — GitHub requires a User-Agent; use token if available from your config (optional)
$headers = @{
    'User-Agent' = 'InstallationScripts/1.0'
    'Accept'     = 'application/vnd.github.v3+json'
}

if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden — verwende authentifizierte Anfrage." -Level "DEBUG"
}

$latestRelease = $null

try {
    # Use -Headers and -ErrorAction Stop so we can catch non-2xx responses
    $latestRelease = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub API Response erhalten: Exists = $($null -ne $latestRelease)" -Level "DEBUG"
} catch {
    # Try to extract a helpful message from the error response (if present)
    $errorBody = $null
    try {
        if ($_.Exception.Response -ne $null) {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $errorBody = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
        }
    } catch {
        # ignore any errors while trying to read error body
        $errorBody = $null
    }

    if ($errorBody) {
        # Try to convert JSON error body to object and inspect message field
        try {
            $json = $errorBody | ConvertFrom-Json -ErrorAction Stop
            $apiMessage = $json.message
        } catch {
            $apiMessage = $errorBody
        }
    } else {
        $apiMessage = $_.Exception.Message
    }

    # Detect rate-limit / quota messages (simple case-insensitive search)
    if ($apiMessage -and ($apiMessage -match '(rate limit|rate_limit|API rate limit|rate limit exceeded)')) {
        Write_LogEntry -Message "GitHub API Rate-Limit erkannt: $($apiMessage)" -Level "WARNING"
        Write_LogEntry -Message "Hinweis: Unauthentifizierte Anfragen haben eine niedrige Grenze. Mit einem GitHub-Token (in PowerShellVariables) erhöht sich das Limit." -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Fehler beim Abruf der GitHub API: $($apiMessage)" -Level "ERROR"
    }

    # Do not rethrow — set $latestRelease to $null and continue script flow
    $latestRelease = $null
}

if ($latestRelease) {
    $onlineVersion = $latestRelease.tag_name.TrimStart("v")
    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $onlineVersion" -foregroundcolor "Cyan"
    Write-Host ""
    Write_LogEntry -Message "Online Version ermittelt: $($onlineVersion); Lokale Version: $($localVersion)" -Level "INFO"
	
    # Compare the local and online file versions
    if ($onlineVersion -gt $localVersion) {
        Write_LogEntry -Message "Online Version ist neuer: Online $($onlineVersion) > Lokal $($localVersion)" -Level "INFO"
        # Get the download URL and filename from the release assets
        $downloadAsset = $latestRelease.assets | Where-Object { $_.name -like "*-win-x64.msi" }
        if ($downloadAsset) {
            $downloadUrl = $downloadAsset.browser_download_url
            $downloadFilename = $downloadAsset.name
            $downloadPath = "$InstallationFolder\$downloadFilename"
            Write_LogEntry -Message "Download-Asset gefunden: Name=$($downloadFilename); URL=$($downloadUrl); Zielpfad=$($downloadPath)" -Level "DEBUG"

            # Download the updated installer with the same filename as online
            try {
                Write_LogEntry -Message "Starte Download: $($downloadUrl) -> $($downloadPath)" -Level "INFO"
                $webClient = New-Object System.Net.WebClient
                [void](Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath)
                $webClient.Dispose()
                Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Download von $($downloadUrl): $($($_.Exception.Message))" -Level "ERROR"
            }

    		# Check if the file was completely downloaded
    		if (Test-Path $downloadPath) {
    			# Remove the old installer
                try {
        			Remove-Item -Path $localFile.FullName -Force
                    Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFile.FullName)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($localFile.FullName): $($($_.Exception.Message))" -Level "WARNING"
                }

    			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                Write_LogEntry -Message "$($ProgramName) Update erfolgreich: $($downloadFilename)" -Level "SUCCESS"
    		} else {
    			Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
                Write_LogEntry -Message "Download fehlgeschlagen; Datei nicht gefunden: $($downloadPath)" -Level "ERROR"
    		}
        } else {
            Write_LogEntry -Message "Kein Download-Asset mit '*-win-x64.msi' gefunden im Release." -Level "WARNING"
        }
    } else {
		Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Online Update verfügbar. Online: $($onlineVersion); Lokal: $($localVersion)" -Level "INFO"
    }
} else {
    # latestRelease == $null -> we either hit rate-limit or an API error; log already created above
    Write_LogEntry -Message "Keine Online-Release-Information verfügbar; überspringe Online-Vergleich." -Level "WARNING"
}

Write-Host ""

#Check Installed Version / Install if neded
$FoundFile = Get-ChildItem -Path $localFilePath | Select-Object -Last 1
if ($FoundFile) {
    $InstallationFileName = $FoundFile.Name
    $localVersion = ($InstallationFileName -replace 'PowerShell-(\d+\.\d+\.\d+).*', '$1')
    Write_LogEntry -Message "Erneut lokale Installationsdatei bestimmt: $($FoundFile.FullName); Version: $($localVersion)" -Level "DEBUG"
} else {
    $InstallationFileName = $null
    $localVersion = "0.0.0"
    Write_LogEntry -Message "Keine lokale Installationsdatei gefunden bei erneutem Check: $($localFilePath)" -Level "WARNING"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Suche: $($RegistryPaths -join '; ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registrypfad gefunden: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
    } else {
        Write_LogEntry -Message "Registrypfad nicht vorhanden: $($RegPath)" -Level "DEBUG"
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
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
	Write_LogEntry -Message "$($ProgramName) ist in Registry gefunden; InstallierteVersion: $($installedVersion); InstallationsdateiVersion: $($localVersion)" -Level "INFO"
	
    if ([version]$installedVersion -lt [version]$localVersion) {
        Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
		$Install = $true
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (Update erforderlich)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
		$Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (Version aktuell)" -Level "DEBUG"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
		$Install = $false
        Write_LogEntry -Message "Installationsentscheidung: Install = $($Install) (installierte Version neuer)" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
	$Install = $false
    Write_LogEntry -Message "$($ProgramName) ist nicht in der Registry gefunden worden (nicht installiert)." -Level "INFO"
}
Write-Host ""
Write_LogEntry -Message "Installationsprüfung abgeschlossen. Install variable: $($Install)" -Level "DEBUG"

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "InstallationFlag gesetzt. Starte Installationsskript mit Flag: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1" -PassInstallationFlag
    Write_LogEntry -Message "Installationsskript mit Flag aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte Installationsskript (Update) ohne Flag: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1" -Level "INFO"
	Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1"
    Write_LogEntry -Message "Installationsskript aufgerufen: $($Serverip)\Daten\Prog\InstallationScripts\Installation\Powershell7Install.ps1" -Level "DEBUG"
}

Write-Host ""
Write-Host "Ändere Powershell.exe zu Pwsh.exe in allen Tasks, wenn PS7 installiert ist." -foregroundcolor "magenta"
Write_LogEntry -Message "Starte Aufgabenplanungs-Update-Skript: $($Serverip)\Daten\Customize_Windows\Scripte\Aufgabenplannung_powershell_to_pwsh.ps1" -Level "INFO"
Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath "$Serverip\Daten\Customize_Windows\Scripte\Aufgabenplannung_powershell_to_pwsh.ps1"
Write_LogEntry -Message "Aufgabenplanungs-Update-Skript aufgerufen: $($Serverip)\Daten\Customize_Windows\Scripte\Aufgabenplannung_powershell_to_pwsh.ps1" -Level "DEBUG"

Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht." -Level "INFO"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
