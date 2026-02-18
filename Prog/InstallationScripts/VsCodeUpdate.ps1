param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Microsoft Visual Studio Code"
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
    exit 1
}

$localFilePathPattern = Join-Path -Path $InstallationFolder -ChildPath "VSCodeUserSetup-x64-*.exe"
Write_LogEntry -Message "Suche lokale Installer mit Muster: $($localFilePathPattern)" -Level "DEBUG"

# Get the local file version (if file exists)
$localFile = Get-ChildItem -Path $localFilePathPattern -ErrorAction SilentlyContinue | Select-Object -Last 1
if ($localFile) {
    try {
        $localVersion = (Get-ItemProperty -Path $localFile.FullName -ErrorAction Stop).VersionInfo.ProductVersion
        Write_LogEntry -Message "Gefundene lokale Datei: $($localFile.FullName); Version: $($localVersion)" -Level "DEBUG"
    } catch {
        $localVersion = $null
        Write_LogEntry -Message ("Fehler beim Lesen der Dateiversion der lokalen Datei {0}: {1}" -f $localFile.FullName, $_.Exception.Message) -Level "WARNING"
    }
} else {
    $localVersion = $null
    Write_LogEntry -Message "Keine lokale Installationsdatei gefunden mit Muster: $($localFilePathPattern)" -Level "INFO"
}

# Prepare GitHub API call with optional token
$apiUrl = "https://api.github.com/repos/microsoft/vscode/releases/latest"
Write_LogEntry -Message "Rufe GitHub API ab: $($apiUrl)" -Level "DEBUG"

# Ensure TLS1.2
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Prepare headers (GitHub requires User-Agent). Use $GithubToken if provided in your PowerShellVariables.ps1
$headers = @{
    'User-Agent' = 'InstallationScripts/1.0'
    'Accept'     = 'application/vnd.github.v3+json'
}

if ($GithubToken) {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token aus config vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
}

# Call GitHub API with robust error handling
$latestRelease = $null
try {
    Write_LogEntry -Message "Sende Invoke-RestMethod an GitHub API (Headers vorhanden: $($headers.Keys -join ', '))" -Level "DEBUG"
    $latestRelease = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
    Write_LogEntry -Message "GitHub API erfolgreich abgefragt; TagName: $($latestRelease.tag_name)" -Level "DEBUG"
} catch {
    # try to parse response message (rate limit etc.)
    $err = $_.Exception
    $msg = $err.Message
    # If webexception with response body, try to read it
    if ($err -is [System.Net.WebException] -and $err.Response) {
        try {
            $stream = $err.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $body = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            if ($body) {
                try {
                    $jsonErr = $body | ConvertFrom-Json -ErrorAction Stop
                    if ($jsonErr.message) { $msg = $jsonErr.message }
                } catch {
                    # leave original message
                    $msg = $body
                }
            }
        } catch {
            # ignore
        }
    }

    if ($msg -match '(rate limit|rate_limit|rate limit exceeded|API rate limit|403)') {
        Write_LogEntry -Message ("GitHub API Rate-Limit / Zugriff verweigert erkannt: {0}" -f $msg) -Level "WARNING"
    } else {
        Write_LogEntry -Message ("Fehler bei GitHub API-Abfrage: {0}" -f $msg) -Level "ERROR"
    }
    $latestRelease = $null
}

# If we have a release, extract version and compare
if ($latestRelease) {
    try {
        $latestVersion = $latestRelease.tag_name.TrimStart('v')
        Write_LogEntry -Message "Ermittelte Online-Version: $($latestVersion)" -Level "INFO"
    } catch {
        $latestVersion = $null
        Write_LogEntry -Message ("Fehler beim Extrahieren der Online-Version: {0}" -f $_.Exception.Message) -Level "ERROR"
    }
} else {
    $latestVersion = $null
}

Write-Host ""
Write-Host "Lokale Version: $($localVersion)" -foregroundcolor "Cyan"
Write-Host "Online Version: $($latestVersion)" -foregroundcolor "Cyan"
Write-Host ""

# Compare and download if necessary
if ($latestVersion) {
    $needUpdate = $false
    try {
        if ($localVersion) {
            $needUpdate = ([version]($latestVersion) -gt [version]($localVersion))
        } else {
            # no local file -> treat as update available
            $needUpdate = $true
        }
    } catch {
        # If version parsing fails, be conservative and skip update check
        Write_LogEntry -Message ("Fehler beim Vergleichen von Versionen: {0}" -f $_.Exception.Message) -Level "WARNING"
        $needUpdate = $false
    }

    if ($needUpdate) {
        Write_LogEntry -Message "Update erkannt: Local $($localVersion) -> Online $($latestVersion)" -Level "INFO"

        # Build download link and destination
        # VS Code direct update URL pattern (stable): https://update.code.visualstudio.com/<version>/win32-x64-user/stable
        $downloadLink = "https://update.code.visualstudio.com/$latestVersion/win32-x64-user/stable"

        if ($localFile) {
            $downloadDir = $localFile.Directory.FullName
        } else {
            $downloadDir = $InstallationFolder
            # ensure dir exists
            if (-not (Test-Path -Path $downloadDir)) {
                try { New-Item -Path $downloadDir -ItemType Directory -Force | Out-Null } catch { Write_LogEntry -Message ("Konnte Zielverzeichnis nicht erstellen: {0}" -f $downloadDir) -Level "ERROR" }
            }
        }
        $downloadPath = Join-Path -Path $downloadDir -ChildPath ("VSCodeUserSetup-x64-$latestVersion.exe")
        Write_LogEntry -Message "Download-Link konstruiert: $($downloadLink). Zielpfad: $($downloadPath)" -Level "DEBUG"

        # Download via WebClient (fast) with optional Authorization header usage
        $webClient = $null
        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("user-agent", $headers['User-Agent'])
            if ($headers.ContainsKey('Authorization')) {
                # WebClient accepts Authorization header
                $webClient.Headers.Add("Authorization", $headers['Authorization'])
            }
            Write_LogEntry -Message "Starte Download: $($downloadLink) -> $($downloadPath)" -Level "INFO"
            $webClient.DownloadFile($downloadLink, $downloadPath)
            Write_LogEntry -Message "Download abgeschlossen: $($downloadPath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message ("Fehler beim Herunterladen $downloadLink : {0}" -f $_.Exception.Message) -Level "ERROR"
            $downloadPath = $null
        } finally {
            if ($webClient) { $webClient.Dispose() }
        }

        # If downloaded, remove old file (if any) and replace
        if ($downloadPath -and (Test-Path -Path $downloadPath)) {
            if ($localFile) {
                try {
                    Remove-Item -Path $localFile.FullName -Force -ErrorAction SilentlyContinue
                    Write_LogEntry -Message "Alte Installationsdatei entfernt: $($localFile.FullName)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message ("Warnung: Konnte alte Datei nicht entfernen: {0}" -f $_.Exception.Message) -Level "WARNING"
                }
            }
            Write-Host "$($ProgramName) wurde aktualisiert.." -foregroundcolor "Green"
            Write_LogEntry -Message "$($ProgramName) wurde aktualisiert; neue Datei: $($downloadPath)" -Level "SUCCESS"
        } else {
            Write_LogEntry -Message "Download nicht vorhanden oder fehlgeschlagen; Update abgebrochen." -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Kein Online Update verfügbar oder lokale Version ist aktuell." -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $($ProgramName) ist aktuell." -foregroundcolor "DarkGray"
    }
} else {
    Write_LogEntry -Message "Keine Online-Version ermittelt; Online-Check übersprungen." -Level "WARNING"
}

Write-Host ""

# --- Nachprüfung / Install-Flag handling ---
# Re-evaluate local file for installation checks
$localFile = Get-ChildItem -Path $localFilePathPattern -ErrorAction SilentlyContinue | Select-Object -Last 1
if ($localFile) {
    try {
        $localVersion = (Get-ItemProperty -Path $localFile.FullName -ErrorAction Stop).VersionInfo.ProductVersion
        Write_LogEntry -Message "Für Prüfungen gefundene lokale Datei: $($localFile.FullName); Version: $($localVersion)" -Level "DEBUG"
    } catch {
        $localVersion = $null
        Write_LogEntry -Message ("Fehler beim Lesen der Produktversion der lokalen Datei {0}: {1}" -f $localFile.FullName, $_.Exception.Message) -Level "DEBUG"
    }
} else {
    $localVersion = $null
    Write_LogEntry -Message "Keine lokale Datei für nachfolgende Prüfungen gefunden" -Level "DEBUG"
}

# Registry check for installed VS Code
$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade gesetzt: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Durchsuche Uninstall-Pfad: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "Microsoft Visual Studio Code*" }
    } else {
        Write_LogEntry -Message "Uninstall-Pfad existiert nicht: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$($ProgramName) ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$($ProgramName) in Registry gefunden; InstalledVersion=$($installedVersion); LocalFileVersion=$($localVersion)" -Level "INFO"

    try {
        if ($localVersion -and $installedVersion -and ([version]$installedVersion -lt [version]$localVersion)) {
            $Install = $true
            Write_LogEntry -Message "Install wird gesetzt: $($Install) (Update erforderlich)" -Level "INFO"
        } else {
            $Install = $false
            Write_LogEntry -Message "Install = $($Install) (keine Aktion erforderlich oder Vergleich nicht möglich)" -Level "DEBUG"
        }
    } catch {
        $Install = $false
        Write_LogEntry -Message ("Fehler beim Vergleichen der installierten und lokalen Version: {0}" -f $_.Exception.Message) -Level "WARNING"
    }
} else {
    $Install = $false
    Write_LogEntry -Message "$($ProgramName) nicht in Registry gefunden; Install = $($Install)" -Level "INFO"
}

Write-Host ""

#Install if needed
if($InstallationFlag){
    Write_LogEntry -Message "Starte externes Installationsscript (InstallationFlag) via $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\VsCodeInstall.ps1" `
		-InstallationFlag
    Write_LogEntry -Message "Rückkehr von VsCodeInstall.ps1 nach InstallationFlag-Aufruf" -Level "DEBUG"
} elseif ($Install -eq $true) {
    Write_LogEntry -Message "Starte externes Installationsscript (Install) via $($PSHostPath)" -Level "INFO"
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Prog\InstallationScripts\Installation\VsCodeInstall.ps1"
    Write_LogEntry -Message "Rückkehr von VsCodeInstall.ps1 nach Install-Aufruf" -Level "DEBUG"
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