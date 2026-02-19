param(
    [switch]$InstallationFlagX86 = $false,
    [switch]$InstallationFlagX64 = $false
)

$ProgramName = "VC Redist x86"
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

Write_LogEntry -Message "Script gestartet mit InstallationFlagX86: $($InstallationFlagX86); InstallationFlagX64: $($InstallationFlagX64)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt auf: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

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

# === TLS / Headers für Webzugriffe ===
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
$headers = @{ 'User-Agent' = 'InstallationScripts/1.0'; 'Accept' = 'text/plain,application/octet-stream' }

if ($null -ne $GithubToken -and $GithubToken -ne '') {
    $headers['Authorization'] = "token $GithubToken"
    Write_LogEntry -Message "GitHub Token vorhanden - verwende authentifizierte Anfrage." -Level "DEBUG"
} else {
    Write_LogEntry -Message "Kein GitHub Token; benutze anonyme Abfrage (Rate-Limits möglich)." -Level "DEBUG"
}

$localFilePathX86 = Join-Path -Path $InstallationFolder -ChildPath "AutoIt_Scripts\VC_redist.x86.exe"
$localFilePathX64 = Join-Path -Path $InstallationFolder -ChildPath "VirtualBox\VC_redist.x64.exe"
Write_LogEntry -Message "Local File Path X86: $($localFilePathX86); Local File Path X64: $($localFilePathX64)" -Level "DEBUG"

# sichere Ermittlung der lokalen Versionen (falls Datei fehlt -> 0.0.0)
function Get-LocalProductVersion {
    param([string]$Path)
    try {
        if ($Path -and (Test-Path -Path $Path)) {
            $v = (Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.ProductVersion
            if ($v) { return $v }
        }
    } catch {
        Write_LogEntry -Message ("Fehler beim Lesen der Produktversion für {0}: {1}" -f $Path, $_.Exception.Message) -Level "DEBUG"
    }
    return "0.0.0"
}

$localFileVersionX86 = Get-LocalProductVersion -Path $localFilePathX86
$localFileVersionX64 = Get-LocalProductVersion -Path $localFilePathX64
Write_LogEntry -Message "Lokale Versionen ermittelt: X86=$($localFileVersionX86); X64=$($localFileVersionX64)" -Level "DEBUG"

# ---- Funktion zum Abrufen der neuesten VC Redist Version ----
function Get-LatestVCRedistVersion {
    try {
        # Methode 1: GitHub Gist - zuverlässige Community-Quelle
        $gistUrl = "https://gist.githubusercontent.com/ChuckMichael/7366c38f27e524add3c54f710678c98b/raw/377c255f48319891068d29d6e4588ef5bc378a4e/vcredistr.md"
        Write_LogEntry -Message "Versuche Versions-Info von GitHub Gist abzurufen..." -Level "DEBUG"
        
        $response = Invoke-RestMethod -Uri $gistUrl -Headers $headers -ErrorAction Stop
        
        # Suche nach Versionsnummer im Format (14.xx.xxxxx) in Klammern
        if ($response -match '\((\d+\.\d+\.\d+)\)') {
            $version = $matches[1]
            Write_LogEntry -Message "Gefundene Version von GitHub Gist: $version" -Level "INFO"
            
            # Download-Links extrahieren (falls vorhanden)
            if ($response -match 'https://aka\.ms/vc\d+/vc_redist\.x64\.exe') {
                $script:downloadLinkX64 = $matches[0]
            }
            if ($response -match 'https://aka\.ms/vc\d+/vc_redist\.x86\.exe') {
                $script:downloadLinkX86 = $matches[0]
            }
            
            return $version
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen von GitHub Gist: $($_.Exception.Message)" -Level "WARNING"
    }
    
    try {
        # Methode 2: Filehorse.com als Fallback
        $url = "https://www.filehorse.com/download-microsoft-visual-c-redistributable-package-64/"
        Write_LogEntry -Message "Versuche Versions-Info von Filehorse abzurufen..." -Level "DEBUG"
        
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
        
        # Suche nach Versionsnummer im Format 14.xx.xxxxx.x
        if ($response.Content -match '14\.\d+\.\d+\.\d+') {
            $version = $matches[0]
            Write_LogEntry -Message "Gefundene Version von Filehorse: $version" -Level "INFO"
            
            # Standard Microsoft Download-Links verwenden
            $script:downloadLinkX64 = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
            $script:downloadLinkX86 = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
            
            return $version
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen von Filehorse: $($_.Exception.Message)" -Level "WARNING"
    }
    
    Write_LogEntry -Message "Konnte keine Online-Version ermitteln" -Level "ERROR"
    return $null
}

# ---- Abrufen der neuesten Version ----
Write_LogEntry -Message "Rufe Versionsinformationen ab..." -Level "DEBUG"

$latestVersion = Get-LatestVCRedistVersion
$downloadLinkX64 = $null
$downloadLinkX86 = $null

if ($latestVersion) {
    Write_LogEntry -Message "Ermittelte neueste Version: $latestVersion" -Level "INFO"
    
    # Falls nicht durch Gist gesetzt, verwende Standard-Links
    if (-not $downloadLinkX64) {
        $downloadLinkX64 = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    }
    if (-not $downloadLinkX86) {
        $downloadLinkX86 = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
    }
    
    Write_LogEntry -Message "Download-Link X64: $downloadLinkX64" -Level "DEBUG"
    Write_LogEntry -Message "Download-Link X86: $downloadLinkX86" -Level "DEBUG"
    
    Write_LogEntry -Message "Prüfe Updates für VC Redist - Online: $($latestVersion); Lokal X86: $($localFileVersionX86); Lokal X64: $($localFileVersionX64)" -Level "INFO"
    Write-Host ""
    Write-Host "Lokale Version X86: $localFileVersionX86" -ForegroundColor Cyan
    Write-Host "Lokale Version X64: $localFileVersionX64" -ForegroundColor Cyan
    Write-Host "Online Version: $latestVersion" -ForegroundColor Cyan
    Write-Host ""

    function VersionLess($a, $b) {
        try { 
            $verA = [version]($a -replace '[^\d\.]','')
            $verB = [version]($b -replace '[^\d\.]','')
            return $verA -lt $verB
        } catch { 
            return $true 
        }
    }

    # --- X86 update ---
    if ($downloadLinkX86 -and (VersionLess $localFileVersionX86 $latestVersion -or $InstallationFlagX86)) {
        $downloadPathX86 = Join-Path -Path $env:TEMP -ChildPath "VC_redist.x86.exe"
        Write_LogEntry -Message "Starte Download X86: $($downloadLinkX86) -> $($downloadPathX86)" -Level "INFO"
        $wc = $null
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add("user-agent", $headers['User-Agent'])
            if ($headers.ContainsKey('Authorization')) { $wc.Headers.Add("Authorization", $headers['Authorization']) }
            $wc.DownloadFile($downloadLinkX86, $downloadPathX86)
            Write_LogEntry -Message "Download abgeschlossen X86: $($downloadPathX86)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message ("Fehler beim Herunterladen X86 {0}: {1}" -f $downloadLinkX86, $_.Exception.Message) -Level "ERROR"
            $downloadPathX86 = $null
        } finally {
            if ($wc) { $wc.Dispose() }
        }

        if ($downloadPathX86 -and (Test-Path $downloadPathX86)) {
            $destDir = Split-Path -Path $localFilePathX86 -Parent
            if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
            try { if (Test-Path $localFilePathX86) { Remove-Item -Path $localFilePathX86 -Force -ErrorAction SilentlyContinue } } catch { Write_LogEntry -Message ("Warnung: altes X86 nicht lösbar: {0}" -f $_.Exception.Message) -Level "WARNING" }
            try {
                Move-Item -Path $downloadPathX86 -Destination $localFilePathX86 -Force
                Write_LogEntry -Message "Neue X86-Installationsdatei verschoben nach: $($localFilePathX86)" -Level "SUCCESS"
                Write-Host "VC Redist x86 wurde aktualisiert." -ForegroundColor Green
            } catch {
                Write_LogEntry -Message ("Fehler beim Verschieben der X86-Datei: {0}" -f $_.Exception.Message) -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Download X86 fehlgeschlagen oder Datei nicht gefunden: $($downloadPathX86)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Kein Update für X86 erforderlich oder DownloadLink fehlt." -Level "DEBUG"
    }

    # --- X64 update ---
    if ($downloadLinkX64 -and (VersionLess $localFileVersionX64 $latestVersion -or $InstallationFlagX64)) {
        $downloadPathX64 = Join-Path -Path $env:TEMP -ChildPath "VC_redist.x64.exe"
        Write_LogEntry -Message "Starte Download X64: $($downloadLinkX64) -> $($downloadPathX64)" -Level "INFO"
        $wc2 = $null
        try {
            $wc2 = New-Object System.Net.WebClient
            $wc2.Headers.Add("user-agent", $headers['User-Agent'])
            if ($headers.ContainsKey('Authorization')) { $wc2.Headers.Add("Authorization", $headers['Authorization']) }
            $wc2.DownloadFile($downloadLinkX64, $downloadPathX64)
            Write_LogEntry -Message "Download abgeschlossen X64: $($downloadPathX64)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message ("Fehler beim Herunterladen X64 {0}: {1}" -f $downloadLinkX64, $_.Exception.Message) -Level "ERROR"
            $downloadPathX64 = $null
        } finally {
            if ($wc2) { $wc2.Dispose() }
        }

        if ($downloadPathX64 -and (Test-Path $downloadPathX64)) {
            $destDir64 = Split-Path -Path $localFilePathX64 -Parent
            if (-not (Test-Path $destDir64)) { New-Item -Path $destDir64 -ItemType Directory -Force | Out-Null }
            try { if (Test-Path $localFilePathX64) { Remove-Item -Path $localFilePathX64 -Force -ErrorAction SilentlyContinue } } catch { Write_LogEntry -Message ("Warnung: altes X64 nicht lösbar: {0}" -f $_.Exception.Message) -Level "WARNING" }
            try {
                Move-Item -Path $downloadPathX64 -Destination $localFilePathX64 -Force
                Write_LogEntry -Message "Neue X64-Installationsdatei verschoben nach: $($localFilePathX64)" -Level "SUCCESS"
                Write-Host "VC Redist x64 wurde aktualisiert." -ForegroundColor Green
            } catch {
                Write_LogEntry -Message ("Fehler beim Verschieben der X64-Datei: {0}" -f $_.Exception.Message) -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Download X64 fehlgeschlagen oder Datei nicht gefunden: $($downloadPathX64)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Kein Update für X64 erforderlich oder DownloadLink fehlt." -Level "DEBUG"
    }
} else {
    Write_LogEntry -Message "Keine Online-Version ermittelt; überspringe Download-/Update-Checks." -Level "WARNING"
}

# ---- Registry / Installcheck und tatsächliche Installation ----
try {
    $ProgramName = "VC Redist x86"
    $localVersion = Get-LocalProductVersion -Path $localFilePathX86
    Write_LogEntry -Message "Für $ProgramName ermittelte lokale Versionsdatei: $($localFilePathX86); Version: $($localVersion)" -Level "DEBUG"
} catch {
    $localVersion = "0.0.0"
    Write_LogEntry -Message ("Fehler beim Lesen lokaler Version für x86: {0}" -f $_.Exception.Message) -Level "DEBUG"
}

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Registry-Pfade für Suche: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Durchsuche Registry-Pfad: $($RegPath) nach Redistributable (x86)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like '*-2022 Redistributable (x86)*' }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht vorhanden: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path -and $Path.Count -gt 0) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$ProgramName in Registry gefunden; InstalledVersion=$($installedVersion); LocalFileVersion=$($localVersion)" -Level "INFO"

    try {
        if ([version]($installedVersion -replace '[^\d\.]','') -lt [version]($localVersion -replace '[^\d\.]','')) {
            $Installx86 = $true
            Write_LogEntry -Message "Installx86 gesetzt auf $($Installx86)" -Level "INFO"
        } elseif ([version]($installedVersion -replace '[^\d\.]','') -eq [version]($localVersion -replace '[^\d\.]','')) {
            $Installx86 = $false
            Write_LogEntry -Message "Installx86 gesetzt auf $($Installx86)" -Level "DEBUG"
        } else {
            $Installx86 = $false
            Write_LogEntry -Message "Installx86 gesetzt auf $($Installx86) (installierte Version höher oder Vergleich nicht möglich)" -Level "WARNING"
        }
    } catch {
        $Installx86 = $false
        Write_LogEntry -Message ("Fehler beim Vergleichen von Versionen (x86): {0}" -f $_.Exception.Message) -Level "WARNING"
    }
} else {
    $Installx86 = $false
    Write_LogEntry -Message "VC Redist x86 nicht in Registry gefunden; Installx86=$($Installx86)" -Level "INFO"
}

# x64 registry check
try {
    $ProgramName = "VC Redist x64"
    $localVersion = Get-LocalProductVersion -Path $localFilePathX64
    Write_LogEntry -Message "Für $ProgramName ermittelte lokale Versionsdatei: $($localFilePathX64); Version: $($localVersion)" -Level "DEBUG"
} catch {
    $localVersion = "0.0.0"
    Write_LogEntry -Message ("Fehler beim Lesen lokaler Version für x64: {0}" -f $_.Exception.Message) -Level "DEBUG"
}

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Durchsuche Registry-Pfad: $($RegPath) nach Redistributable (x64)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like '*-2022 Redistributable (x64)*' }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht vorhanden: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path -and $Path.Count -gt 0) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
    Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
    Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
    Write_LogEntry -Message "$ProgramName in Registry gefunden; InstalledVersion=$($installedVersion); LocalFileVersion=$($localVersion)" -Level "INFO"

    try {
        if ([version]($installedVersion -replace '[^\d\.]','') -lt [version]($localVersion -replace '[^\d\.]','')) {
            $Installx64 = $true
            Write_LogEntry -Message "Installx64 gesetzt auf $($Installx64)" -Level "INFO"
        } elseif ([version]($installedVersion -replace '[^\d\.]','') -eq [version]($localVersion -replace '[^\d\.]','')) {
            $Installx64 = $false
            Write_LogEntry -Message "Installx64 gesetzt auf $($Installx64)" -Level "DEBUG"
        } else {
            $Installx64 = $false
            Write_LogEntry -Message "Installx64 gesetzt auf $($Installx64) (installierte Version höher oder Vergleich nicht möglich)" -Level "WARNING"
        }
    } catch {
        $Installx64 = $false
        Write_LogEntry -Message ("Fehler beim Vergleichen von Versionen (x64): {0}" -f $_.Exception.Message) -Level "WARNING"
    }
} else {
    $Installx64 = $false
    Write_LogEntry -Message "VC Redist x64 nicht in Registry gefunden; Installx64=$($Installx64)" -Level "INFO"
}

Write-Host ""
Write_LogEntry -Message "Installationsentscheidung: Installx86=$($Installx86); Installx64=$($Installx64); Flags: X86=$($InstallationFlagX86); X64=$($InstallationFlagX64)" -Level "DEBUG"

# Install if needed or if flags set
if (($Installx86 -eq $true) -or $InstallationFlagX86) {
    Write-Host "Microsoft Visual C++ x86 wird installiert" -ForegroundColor Magenta
    Write_LogEntry -Message "Starte Installation VC Redist x86 (Installx86: $($Installx86); Flag: $($InstallationFlagX86))" -Level "INFO"

    $vcInstaller = Get-ChildItem -Path (Join-Path -Path $Serverip -ChildPath "Daten\Prog\AutoIt_Scripts\VC_redist*.exe") -ErrorAction SilentlyContinue | Select-Object -First 1
    Write_LogEntry -Message "Gefundene VC x86 Installer: $($vcInstaller)" -Level "DEBUG"
    if ($vcInstaller) {
        try {
            [void](Invoke-InstallerFile -FilePath $vcInstaller.FullName -Arguments '/install','/passive','/qn','/norestart' -Wait)
            Write_LogEntry -Message "Prozess für VC x86 Installer abgeschlossen: $($vcInstaller.FullName)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message ("Fehler beim Starten des VC x86 Installers {0}: {1}" -f $vcInstaller.FullName, $_.Exception.Message) -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Kein VC x86 Installer auf Server gefunden; überspringe." -Level "ERROR"
    }
}

if (($Installx64 -eq $true) -or $InstallationFlagX64) {
    Write-Host "Microsoft Visual C++ x64 wird installiert" -ForegroundColor Magenta
    Write_LogEntry -Message "Starte Installation VC Redist x64 (Installx64: $($Installx64); Flag: $($InstallationFlagX64))" -Level "INFO"

    $vcInstaller64 = Get-ChildItem -Path (Join-Path -Path $Serverip -ChildPath "Daten\Prog\VirtualBox\VC*.exe") -ErrorAction SilentlyContinue | Select-Object -First 1
    Write_LogEntry -Message "Gefundene VC x64 Installer: $($vcInstaller64)" -Level "DEBUG"
    if ($vcInstaller64) {
        try {
            [void](Invoke-InstallerFile -FilePath $vcInstaller64.FullName -Arguments '/install','/passive','/qn','/norestart' -Wait)
            Write_LogEntry -Message "Prozess für VC x64 Installer abgeschlossen: $($vcInstaller64.FullName)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message ("Fehler beim Starten des VC x64 Installers {0}: {1}" -f $vcInstaller64.FullName, $_.Exception.Message) -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Kein VC x64 Installer auf Server gefunden; überspringe." -Level "ERROR"
    }
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
