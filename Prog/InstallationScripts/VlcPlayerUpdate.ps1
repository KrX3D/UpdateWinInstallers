param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "VLC"
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

# Import DeployToolkit for shared version/install helpers
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) {
    Write_LogEntry -Message "DeployToolkit nicht gefunden: $dtPath" -Level "ERROR"
    exit 1
}
Import-Module -Name $dtPath -Force -ErrorAction Stop

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

# Function to parse the version number from the filename
function ParseVersion {
    param (
        [string]$Filename
    )

    $version = Get-VersionFromFileName -Name $Filename -Regex 'vlc-(\d+\.\d+\.\d+)-win64\.exe'
    if ($version) {
        Write_LogEntry -Message "Version aus Dateiname $($Filename) geparst: $version" -Level "DEBUG"
        return $version.ToString()
    }

    Write_LogEntry -Message "Konnte Version nicht aus Dateiname $($Filename) parsen" -Level "DEBUG"
    return $null
}

function Set-Tls12ForWebClient {
    [CmdletBinding()]
    param()

    try {
        $tls12 = [Net.SecurityProtocolType]::Tls12
        $securityProtocol = [Net.ServicePointManager]::SecurityProtocol
        if (($securityProtocol -band $tls12) -eq 0) {
            [Net.ServicePointManager]::SecurityProtocol = $securityProtocol -bor $tls12
        }

        # Avoid 100-continue delays/timeouts in some environments.
        [Net.ServicePointManager]::Expect100Continue = $false

        # Some environments still require TLS 1.1 or legacy TLS enabled alongside TLS 1.2.
        try {
            $tls11 = [Net.SecurityProtocolType]::Tls11
            if (([Net.ServicePointManager]::SecurityProtocol -band $tls11) -eq 0) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $tls11
            }
        } catch { }

        try {
            $tls10 = [Net.SecurityProtocolType]::Tls
            if (([Net.ServicePointManager]::SecurityProtocol -band $tls10) -eq 0) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $tls10
            }
        } catch { }

        Write_LogEntry -Message "TLS-Protokolle aktiv: $([Net.ServicePointManager]::SecurityProtocol)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Konnte TLS-Protokolle nicht aktivieren: $($_)" -Level "WARNING"
    }
}

# Function to check if a newer version of VLC Player is available and initiate the download
function CheckVLCVersion {
    param (
        [string]$InstalledVersion
    )

    $url = "https://www.videolan.org/vlc/index.html"
    Write_LogEntry -Message "Rufe VLC-Webseite ab: $($url)" -Level "DEBUG"
    $responseContent = Invoke-WebRequestCompat -Uri $url -ReturnContent
    if (-not $responseContent) {
        Write_LogEntry -Message "Konnte VLC-Webseite nicht abrufen: $url" -Level "ERROR"
        return
    }

    # Extract the latest version of VLC Player from the HTML content
    $versionPattern = 'vlc-(\d+\.\d+\.\d+)-win64\.exe'
    $latestVersionObj = Get-OnlineVersionFromContent -Content $responseContent -Regex $versionPattern -SelectLast

    if ($latestVersionObj) {
        $latestVersion = $latestVersionObj.ToString()
        Write_LogEntry -Message "Gefundene Online-Version auf Webseite: $($latestVersion)" -Level "INFO"
    } else {
        Write_LogEntry -Message "Konnte Online-Version auf $($url) nicht ermitteln" -Level "WARNING"
        return
    }
    Write-Host ""
    Write-Host "Lokale Version: $InstalledVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
    Write-Host ""

    if ($latestVersion -gt $InstalledVersion) {
        # Construct the download URL using the latest version
        $downloadUrl = "https://vlc.pixelx.de/vlc/$latestVersion/win64/vlc-$latestVersion-win64.exe"
        Write_LogEntry -Message "Update verfügbar: $($InstalledVersion) -> $($latestVersion). Download-URL: $($downloadUrl)" -Level "INFO"
	
        #Write-Host "Download URL: $downloadUrl"

        $downloadPath = "$InstallationFolder\vlc-$latestVersion-win64.exe"
        Write_LogEntry -Message "Starte Download nach: $($downloadPath)" -Level "INFO"

        Set-Tls12ForWebClient

        #Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath
        $webClient = New-Object System.Net.WebClient
        try {
            $webClient.Headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell WebClient"
            $webClient.Proxy = [System.Net.WebRequest]::DefaultWebProxy
            $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
            $webClient.DownloadFile($downloadUrl, $downloadPath)
            Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Download $($downloadUrl): $($_)" -Level "ERROR"
            Write_LogEntry -Message "Versuche HTTPS erneut ohne Proxy (direkte Verbindung)..." -Level "WARNING"
            try {
                $webClient.Proxy = $null
                $webClient.DownloadFile($downloadUrl, $downloadPath)
                Write_LogEntry -Message "Download ohne Proxy abgeschlossen: $($downloadPath)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "HTTPS ohne Proxy fehlgeschlagen $($downloadUrl): $($_)" -Level "ERROR"
            }
            Write_LogEntry -Message "Versuche HTTP-Mirror mit WebClient (falls HTTPS blockiert ist)..." -Level "WARNING"
            $httpDownloadUrl = "http://download.videolan.org/pub/videolan/vlc/$latestVersion/win64/vlc-$latestVersion-win64.exe"
            Write_LogEntry -Message "HTTP-URL: $($httpDownloadUrl)" -Level "DEBUG"
            try {
                $webClient.Proxy = $null
                $webClient.DownloadFile($httpDownloadUrl, $downloadPath)
                Write_LogEntry -Message "Download über HTTP abgeschlossen: $($downloadPath)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "HTTP-Download fehlgeschlagen $($httpDownloadUrl): $($_)" -Level "ERROR"
            }
        } finally {
            $webClient.Dispose()
        }

        # Check if the file was completely downloaded
        if (Test-Path $downloadPath) {
            # Remove the old installer
            try {
                Remove-Item -Path $vlcPath -Force
                Write_LogEntry -Message "Alte VLC-Installationsdatei entfernt: $($vlcPath)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen der alten Datei $($vlcPath): $($_)" -Level "WARNING"
            }

            Write-Host "$($ProgramName) wurde aktualisiert.." -foregroundcolor "green"
            Write_LogEntry -Message "$($ProgramName) erfolgreich aktualisiert; neue Datei: $($downloadPath)" -Level "SUCCESS"
        } else {
            Write-Host "Download ist fehlgeschlagen. $($ProgramName) wurde nicht aktuallisiert." -foregroundcolor "red"
            Write_LogEntry -Message "Download fehlgeschlagen für $($downloadPath)" -Level "ERROR"
        }

    } else {
        Write-Host "Kein Online Update verfügbar. $($ProgramName) is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Update verfügbar (Local: $($InstalledVersion), Online: $($latestVersion))" -Level "INFO"
    }
}

# Get the latest VLC Player installer file in the directory
$latestInstaller = Get-ChildItem -Path $InstallationFolder -Filter "vlc-*-win64.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($latestInstaller) {
    Write_LogEntry -Message "Gefundene lokale Installer-Datei: $($latestInstaller.FullName)" -Level "DEBUG"
    $vlcPath = $latestInstaller.FullName

    # Get the version from the filename
    $installedVersion = ParseVersion -Filename $vlcPath

    if ($installedVersion) {
        Write_LogEntry -Message "Aufruf CheckVLCVersion mit InstalledVersion: $($installedVersion)" -Level "INFO"
        CheckVLCVersion -InstalledVersion $installedVersion
        Write_LogEntry -Message "Rückkehr aus CheckVLCVersion für Version: $($installedVersion)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Konnte Version aus vorhandener Datei $($vlcPath) nicht parsen" -Level "WARNING"
        #Write-Host "Unable to parse version from the filename: $vlcPath"
    }
} else {
    Write_LogEntry -Message "Kein lokaler VLC-Installer im Ordner $($InstallationFolder) gefunden" -Level "INFO"
    #Write-Host "VLC Player installer not found in the directory: $InstallationFolder"
}

Write-Host ""

#Check Installed Version / Install if neded
$FoundFile = Get-ChildItem -Path $InstallationFolder -Filter "vlc-*-win64.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($FoundFile) {
    Write_LogEntry -Message "Gefundene Installationsdatei für nachfolgende Prüfungen: $($FoundFile.FullName)" -Level "DEBUG"
    $InstallationFileName = $FoundFile.Name
    $localVersion = ParseVersion -Filename $InstallationFileName
} else {
    Write_LogEntry -Message "Keine Installationsdatei für spätere Prüfungen gefunden" -Level "DEBUG"
    $InstallationFileName = $null
    $localVersion = $null
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$installedInfo = Get-InstalledVersionInfo -DisplayNameLike "$($ProgramName)*"

if ($null -ne $installedInfo) {
    $installedVersion = $installedInfo.VersionRaw
    Write-Host "$($ProgramName) ist installiert." -foregroundcolor "green"
    Write-Host "	Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "	Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$($ProgramName) in Registry gefunden; InstalledVersion=$($installedVersion); LocalFileVersion=$($localVersion)" -Level "INFO"

    if (Test-InstallerUpdateRequired -InstalledVersion (ConvertTo-VersionSafe $installedVersion) -InstallerVersion (ConvertTo-VersionSafe $localVersion)) {
        Write-Host "		Veraltete $($ProgramName) ist installiert. Update wird gestartet." -foregroundcolor "magenta"
        $Install = $true
        Write_LogEntry -Message "Install = $($Install) (Update erforderlich)" -Level "INFO"
    } elseif ([version]$installedVersion -eq [version]$localVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
        $Install = $false
        Write_LogEntry -Message "Install = $($Install) (Version aktuell)" -Level "DEBUG"
    } else {
        #Write-Host "$ProgramName is installed, and the installed version ($installedVersion) is higher than the local version ($localVersion)."
        $Install = $false
        Write_LogEntry -Message "Install = $($Install) (installierte Version höher)" -Level "WARNING"
    }
} else {
    #Write-Host "$ProgramName is not installed on this system."
    $Install = $false
    Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden; Install = $($Install)" -Level "INFO"
}
Write-Host ""

#Install if needed
if ($InstallationFlag) {
    Write_LogEntry -Message "Starte externes Installationsscript (InstallationFlag) via $($PSHostPath)" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\VlcPlayerInstall.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Rückkehr von VlcPlayerInstall.ps1 nach InstallationFlag-Aufruf" -Level "DEBUG"
} elseif ($Install -eq $true) {
    Write_LogEntry -Message "Starte externes Installationsscript (Install) via $($PSHostPath)" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\VlcPlayerInstall.ps1"
    Write_LogEntry -Message "Rückkehr von VlcPlayerInstall.ps1 nach Install-Aufruf" -Level "DEBUG"
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
