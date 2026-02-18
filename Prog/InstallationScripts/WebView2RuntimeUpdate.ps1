param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Microsoft Edge Webview 2"
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

# Define the directory path and file wildcard
$InstallationFolder = "$InstallationFolder\ImageGlass"
$fileWildcard = "MicrosoftEdgeWebview2Setup.exe"
Write_LogEntry -Message "InstallationFolder: $($InstallationFolder); FileWildcard: $($fileWildcard)" -Level "DEBUG"

# Get the latest local file path matching the wildcard
#$localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $fileWildcard | Select-Object -Last 1 -ExpandProperty FullName

#Get local Version number
# Define the registry paths and the key to check
$RegistryPaths = @(
    "HKCU:\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
)
$KeyToCheck = "pv"
Write_LogEntry -Message "Registry-Pfade gesetzt: $($RegistryPaths -join ', '); KeyToCheck: $($KeyToCheck)" -Level "DEBUG"

# Initialize an array to store version numbers
$Versions = @()

# Loop through each registry path and check for the key
foreach ($Path in $RegistryPaths) {
    try {
        if (Test-Path $Path) {
            # Get the registry value
            $Version = (Get-ItemProperty -Path $Path -Name $KeyToCheck -ErrorAction SilentlyContinue).$KeyToCheck
            if ($Version) {
                $Versions += $Version
                Write_LogEntry -Message "Gefundene Version in Registry-Pfad $($Path): $($Version)" -Level "DEBUG"
            } else {
                Write_LogEntry -Message "Kein Wert für '$($KeyToCheck)' in Registry-Pfad $($Path) gefunden" -Level "DEBUG"
            }
        } else {
            Write_LogEntry -Message "Registry-Pfad nicht vorhanden: $($Path)" -Level "DEBUG"
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Zugriff auf Registry-Pfad $($Path): $($_)" -Level "ERROR"
        Write-Error "Fehler beim Zugriff auf $Path : $_"
    }
}

# Determine the result
if ($Versions.Count -gt 0) {
    # Get the highest version number
    $localVersion = $Versions | Sort-Object { [Version]$_ } | Select-Object -Last 1
    Write_LogEntry -Message "Ermittelte lokale Version aus Registry: $($localVersion)" -Level "INFO"
} else {
    # Default to 0.0.0.0 if no versions found
    $localVersion = "0.0.0.0"
    Write_LogEntry -Message "Keine lokale Version gefunden; setze Default LocalVersion: $($localVersion)" -Level "WARNING"
}

# Hole die Online-Version
# Get Online Version number
try {
    $WebView2Page = "https://developer.microsoft.com/de-de/microsoft-edge/webview2?form=MA13LH#download"
    Write_LogEntry -Message "Rufe Webseite ab: $($WebView2Page)" -Level "DEBUG"
    $PageContent = Invoke-WebRequest -Uri $WebView2Page -UseBasicParsing
    Write_LogEntry -Message "Webseite abgerufen: $($WebView2Page); ContentLength: $($PageContent.Content.Length)" -Level "DEBUG"

    # Extract the JSON-like data block
    $DataRegex = '"__NUXT_DATA__".*?>(\[.+?\])<'
    if ($PageContent.Content -match $DataRegex) {
        $JsonString = $Matches[1]
        Write_LogEntry -Message "Nuxt JSON-Block extrahiert (Länge: $($JsonString.Length))" -Level "DEBUG"

        # Parse the JSON content into a PowerShell object
        $NuxtDataParsed = $JsonString | ConvertFrom-Json

        # Initialize a list to store all download links with version numbers
        $AllLinks = @()

        # Loop through the NuxtDataParsed and extract download links with version numbers
        for ($i = 0; $i -lt $NuxtDataParsed.Count; $i++) {
			# Check if the item looks like a download link with a version number in it
            if ($NuxtDataParsed[$i] -match 'https:\/\/msedge\.sf\.dl\.delivery\.mp\.microsoft\.com\/filestreamingservice\/files\/[a-f0-9\-]+\/Microsoft\.WebView2\.FixedVersionRuntime\.(\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5})\.x64\.cab') {
            	# Add the link along with the extracted version number
                $Version = $Matches[1]
                $AllLinks += [PSCustomObject]@{
                    Link    = $NuxtDataParsed[$i]
                    Version = $Version
                }
                Write_LogEntry -Message "Gefundener Download-Link mit Version: $($Version)" -Level "DEBUG"
            }
        }
        # If there are any download links found
        if ($AllLinks.Count -gt 0) {
            # Sort by version (descending) and pick the latest
            $LatestLink = $AllLinks | Sort-Object { [Version]$_.Version } -Descending | Select-Object -First 1
            $webVersion = $LatestLink.Version
            Write_LogEntry -Message "Ermittelte Online-Version: $($webVersion)" -Level "INFO"
        } else {
            Write_LogEntry -Message "Es wurden keine Download-Links auf der Seite gefunden." -Level "WARNING"
            Write-Error "Es wurden keine Download-Links auf der Seite gefunden."
        }
    } else {
        Write_LogEntry -Message "Der __NUXT_DATA__ Block konnte nicht extrahiert werden." -Level "ERROR"
        Write-Error "Es konnte der __NUXT_DATA__ Block von der Seite nicht extrahiert werden."
    }
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen oder Verarbeiten der WebView2-Seite: $($_)" -Level "ERROR"
    Write-Error "Es ist ein Fehler aufgetreten: $_"
}

if ($webVersion) {
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $webVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Lokale Version: $($localVersion); Online Version: $($webVersion)" -Level "INFO"
    Write-Host ""
    
    # Compare the local and online versions
    if ([version]$localVersion -lt [version]$webVersion) {
        $newFilePath = Join-Path -Path $InstallationFolder -ChildPath $fileWildcard
        $downloadLink = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
        Write_LogEntry -Message "Update verfügbar: Lokal $($localVersion) < Online $($webVersion). Neuer Pfad: $($newFilePath); DownloadLink: $($downloadLink)" -Level "INFO"
        
        if (Test-Path $newFilePath) {
            Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
            $Install = $true
            Write_LogEntry -Message "Install-Flag gesetzt, Datei bereits vorhanden: $($newFilePath)" -Level "DEBUG"
        } else {
            Write_LogEntry -Message "Starte Download: $($downloadLink) -> $($newFilePath)" -Level "INFO"
            $webClient = New-Object System.Net.WebClient
            try {
                $webClient.DownloadFile($downloadLink, $newFilePath)
                Write_LogEntry -Message "Download abgeschlossen: $($newFilePath)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "Fehler beim Download $($downloadLink): $($_)" -Level "ERROR"
            } finally {
                $webClient.Dispose()
            }
            
            # Check if the file was completely downloaded
            if (Test-Path $newFilePath) {
                Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                Write-Host "		Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
                $Install = $true
                Write_LogEntry -Message "Datei erfolgreich heruntergeladen und Install-Flag gesetzt: $($newFilePath)" -Level "SUCCESS"
            } else {
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktuallisiert." -foregroundcolor "red"
                Write_LogEntry -Message "Download fehlgeschlagen: $($newFilePath)" -Level "ERROR"
            }
        }
    } elseif ([version]$localVersion -eq [version]$webVersion) {
        Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
        $Install = $false
        Write_LogEntry -Message "Keine Aktion erforderlich; lokal gleich online: $($localVersion)" -Level "INFO"
    } else {
        $Install = $false
        Write_LogEntry -Message "Lokale Version ist neuer als Online-Version: Local=$($localVersion); Online=$($webVersion)" -Level "WARNING"
    }
}
Write-Host ""
Write_LogEntry -Message "Install-Flag: $($Install); InstallationFlag: $($InstallationFlag)" -Level "DEBUG"

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "Starte externes Installationsskript mit Flag: $($InstallationFlag)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\WebView2RuntimeInstall.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WebView2RuntimeInstall.ps1 (mit -InstallationFlag)" -Level "DEBUG"
} elseif($Install -eq $true){
    Write_LogEntry -Message "Starte externes Installationsskript (Install=true)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\WebView2RuntimeInstall.ps1"
    Write_LogEntry -Message "Externer Aufruf abgeschlossen: WebView2RuntimeInstall.ps1" -Level "DEBUG"
}

Write-Host ""
Write_LogEntry -Message "Script-Ende erreicht" -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===
