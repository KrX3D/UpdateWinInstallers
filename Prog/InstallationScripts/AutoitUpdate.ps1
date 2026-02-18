param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Autoit"
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
        Initialize_LogSession -ProgramName $ProgramName -ScriptType $ScriptType | Out-Null
    }
}
# === Ende Logger-Header ===

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PSScriptRoot: $($PSScriptRoot)" -Level "DEBUG"

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Versuche Konfigurationsdatei zu laden: $($configPath)" -Level "INFO"

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

# Ensure TLS 1.2 for subsequent requests
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
# Utility: Robust herunterladen (WebClient mit Browser-Header, Invoke-WebRequest, BITS-Fallback)
function Download-File {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$OutFile,
        [string]$Referer = $null
    )

    $ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    $accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"

    try {
        Write_LogEntry -Message "Versuche Download via WebClient: $Url -> $OutFile" -Level "DEBUG"
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent", $ua)
        if ($Referer) { $wc.Headers.Add("Referer", $Referer) }
        $wc.Headers.Add("Accept", $accept)

        $wc.DownloadFile($Url, $OutFile)
        $wc.Dispose()
        Write_LogEntry -Message "Download via WebClient erfolgreich: $OutFile" -Level "SUCCESS"
        return $true
    } catch {
        Write_LogEntry -Message "WebClient-Download fehlgeschlagen: $($_.Exception.Message)" -Level "DEBUG"
        # Fallback: Invoke-WebRequest
        try {
            Write_LogEntry -Message "Versuche Download via Invoke-WebRequest: $Url -> $OutFile" -Level "DEBUG"
            $hdr = @{ "User-Agent" = $ua; "Accept" = $accept }
            if ($Referer) { $hdr["Referer"] = $Referer }
            Invoke-WebRequest -Uri $Url -OutFile $OutFile -Headers $hdr -UseBasicParsing -ErrorAction Stop
            Write_LogEntry -Message "Download via Invoke-WebRequest erfolgreich: $OutFile" -Level "SUCCESS"
            return $true
        } catch {
            Write_LogEntry -Message "Invoke-WebRequest fehlgeschlagen: $($_.Exception.Message)" -Level "DEBUG"
            # Fallback 2: BITS
            try {
                Write_LogEntry -Message "Versuche Download via BITS (Start-BitsTransfer): $Url -> $OutFile" -Level "DEBUG"
                Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
                Write_LogEntry -Message "Download via BITS erfolgreich: $OutFile" -Level "SUCCESS"
                return $true
            } catch {
                Write_LogEntry -Message "Alle Download-Methoden fehlgeschlagen: $($_.Exception.Message)" -Level "ERROR"
                return $false
            }
        }
    }
}

$autoItDownloadUrl = "https://www.autoitscript.com/site/autoit/downloads/"
$sciTEDownloadUrl = "https://www.autoitscript.com/cgi-bin/getfile.pl?../autoit3/scite/download/SciTE4AutoIt3.exe"

Write_LogEntry -Message "AutoIt Download-URL: $($autoItDownloadUrl)" -Level "DEBUG"
Write_LogEntry -Message "SciTE Download-URL: $($sciTEDownloadUrl)" -Level "DEBUG"

$InstallationFolder = "$InstallationFolder\AutoIt_Scripts"
Write_LogEntry -Message "InstallationFolder gesetzt auf: $($InstallationFolder)" -Level "DEBUG"

$localAutoItFile = Get-ChildItem -Path "$InstallationFolder\autoit-v3-setup*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
$localSciTEFile = Get-ChildItem -Path "$InstallationFolder\SciTE4AutoIt3*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1

Write_LogEntry -Message "Lokale AutoIt Datei: $(if ($localAutoItFile) { $localAutoItFile.FullName } else { 'None' })" -Level "DEBUG"
Write_LogEntry -Message "Lokale SciTE Datei: $(if ($localSciTEFile) { $localSciTEFile.FullName } else { 'None' })" -Level "DEBUG"

$localAutoItVersion = $null
$localSciTEVersion = $null
if ($localAutoItFile) {
    try {
        $localAutoItVersion = (Get-ItemProperty -Path $localAutoItFile.FullName).VersionInfo.FileVersion
        Write_LogEntry -Message "Lokale AutoIt Version ermittelt: $($localAutoItVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Dateiinfo der lokalen AutoIt-Datei $($localAutoItFile.FullName): $($_)" -Level "ERROR"
    }
}
if ($localSciTEFile) {
    try {
        $localSciTEVersion = (Get-ItemProperty -Path $localSciTEFile.FullName).VersionInfo.ProductVersion
        Write_LogEntry -Message "Lokale SciTE Version ermittelt: $($localSciTEVersion)" -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen der Dateiinfo der lokalen SciTE-Datei $($localSciTEFile.FullName): $($_)" -Level "ERROR"
    }
}

Write_LogEntry -Message "Rufe AutoIt-Downloadseite ab: $($autoItDownloadUrl)" -Level "INFO"
$autoItPageContent = $null
try {
    $autoItPageContent = Invoke-RestMethod -Uri $autoItDownloadUrl -ErrorAction Stop
    Write_LogEntry -Message "AutoIt Seite abgerufen: Success = $($autoItPageContent -ne $null)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler beim Abrufen der AutoIt-Seite $($autoItDownloadUrl): $($_)" -Level "ERROR"
}

$autoItDownloadLinkPattern = 'v(\d+\.\d+\.\d+\.\d+)'
$autoItMatch = [regex]::Match([string]$autoItPageContent, $autoItDownloadLinkPattern)

$relativeUrlPattern = '(?<=href="\/cgi-bin\/getfile\.pl\?)([^"]+autoit-v3-setup[^"]*)'
$relativeUrlMatch = [regex]::Match([string]$autoItPageContent, $relativeUrlPattern)

if ($autoItMatch.Success -and $relativeUrlMatch.Success) {
    $onlineAutoItVersion = $autoItMatch.Groups[1].Value
    Write_LogEntry -Message "AutoIt Online Version: $($onlineAutoItVersion); Lokale Version: $($localAutoItVersion)" -Level "INFO"
    Write-Host ""
    Write-Host "$ProgramName Lokale Version: $localAutoItVersion" -ForegroundColor "Cyan"
    Write-Host "$ProgramName Online Version: $onlineAutoItVersion" -ForegroundColor "Cyan"
    Write-Host ""

    $relativeUrl = $relativeUrlMatch.Groups[1].Value
    $autoItDownloadLink = "https://www.autoitscript.com/cgi-bin/getfile.pl?$relativeUrl"
    $filename = Split-Path -Path $autoItDownloadLink -Leaf

    Write_LogEntry -Message "Gefundener Download-Link: $($autoItDownloadLink); Filename: $($filename)" -Level "DEBUG"

    if ($onlineAutoItVersion -gt $localAutoItVersion) {
        $autoItSavePath = Join-Path -Path $env:TEMP -ChildPath $filename

        if (Download-File -Url $autoItDownloadLink -OutFile $autoItSavePath -Referer $autoItDownloadUrl) {
            Write_LogEntry -Message "Download abgeschlossen: $($autoItSavePath)" -Level "SUCCESS"
        } else {
            Write_LogEntry -Message "Download ist fehlgeschlagen. $($filename) wurde nicht aktualisiert." -Level "ERROR"
        }

        if (Test-Path $autoItSavePath) {
            Write_LogEntry -Message "Download-Datei vorhanden: $($autoItSavePath)" -Level "DEBUG"

            try {
                if ($autoItSavePath -match '\.zip$') {
                    Expand-Archive -Path $autoItSavePath -DestinationPath $env:TEMP -Force
                    Write_LogEntry -Message "Archiv entpackt nach $($env:TEMP)" -Level "SUCCESS"
                } else {
                    Write_LogEntry -Message "Keine Archiv-Datei (kein .zip). Überspringe Entpacken." -Level "DEBUG"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Entpacken von $($autoItSavePath): $($_)" -Level "ERROR"
            }

            try {
                if ($localAutoItFile) {
                    Write_LogEntry -Message "Entferne alte AutoIt Datei: $($localAutoItFile.FullName)" -Level "DEBUG"
                    Remove-Item -Path $localAutoItFile.FullName -Force
                    Write_LogEntry -Message "Alte AutoIt Datei entfernt: $($localAutoItFile.FullName)" -Level "SUCCESS"
                }

                if (Test-Path $autoItSavePath) {
                    try { Remove-Item -Path $autoItSavePath -Force -ErrorAction SilentlyContinue } catch {}
                    Write_LogEntry -Message "Temporäre Datei (zip/exe) behandelt: $($autoItSavePath)" -Level "DEBUG"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen alter/temporärer Dateien: $($_)" -Level "ERROR"
            }

            try {
                $newAutoItFile = $null
                $cand = Get-ChildItem -Path ("$env:TEMP\autoit*.exe") -ErrorAction SilentlyContinue | Select-Object -Last 1
                if ($cand) { $newAutoItFile = $cand.FullName } elseif (Test-Path $autoItSavePath -and $autoItSavePath -match '\.exe$') { $newAutoItFile = $autoItSavePath }

                if ($newAutoItFile -and (Test-Path $newAutoItFile)) {
                    Write_LogEntry -Message "Verschiebe neue AutoIt Datei $($newAutoItFile) -> $($InstallationFolder)" -Level "DEBUG"
                    Move-Item -Path $newAutoItFile -Destination $InstallationFolder -Force
                    Write_LogEntry -Message "Neue AutoIt Datei verschoben nach $($InstallationFolder)" -Level "SUCCESS"
                } else {
                    Write_LogEntry -Message "Neue AutoIt-Datei nicht gefunden zum Verschieben: $($newAutoItFile)" -Level "ERROR"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Verschieben der neuen AutoIt-Datei: $($_)" -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Download ist fehlgeschlagen. $($filename) wurde nicht aktualisiert." -Level "ERROR"
            Write-Host "Download ist fehlgeschlagen. $filename wurde nicht aktualisiert." -ForegroundColor "Red"
        }

        # Download the latest SciTE for AutoIt
        $sciTESavePath = Join-Path -Path $env:TEMP -ChildPath "SciTE4AutoIt3.exe"
        if (Download-File -Url $sciTEDownloadUrl -OutFile $sciTESavePath -Referer $autoItDownloadUrl) {
            Write_LogEntry -Message "SciTE Download abgeschlossen: $($sciTESavePath)" -Level "SUCCESS"
        } else {
            Write_LogEntry -Message "Download ist fehlgeschlagen. SciTE4AutoIt3.exe wurde nicht aktualisiert." -Level "ERROR"
        }

        if (Test-Path $sciTESavePath) {
            try {
                if ($localSciTEFile) {
                    Write_LogEntry -Message "Entferne alte SciTE Datei: $($localSciTEFile.FullName)" -Level "DEBUG"
                    Remove-Item -Path $localSciTEFile.FullName -Force
                    Write_LogEntry -Message "Alte SciTE Datei entfernt: $($localSciTEFile.FullName)" -Level "SUCCESS"
                }

                $newSciteFile = (Get-ChildItem -Path ("$env:TEMP\SciTE4AutoIt3*.exe") -ErrorAction SilentlyContinue | Select-Object -Last 1).FullName
                if ($newSciteFile -and (Test-Path $newSciteFile)) {
                    Write_LogEntry -Message "Verschiebe neue SciTE Datei $($newSciteFile) -> $($InstallationFolder)" -Level "DEBUG"
                    Move-Item -Path $newSciteFile -Destination $InstallationFolder -Force
                    Write_LogEntry -Message "Neue SciTE Datei verschoben nach $($InstallationFolder)" -Level "SUCCESS"
                } else {
                    Write_LogEntry -Message "Keine neue SciTE Datei gefunden zum Verschieben." -Level "ERROR"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Verarbeiten von SciTE-Dateien: $($_)" -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Download ist fehlgeschlagen. SciTE4AutoIt3.exe wurde nicht aktualisiert." -Level "ERROR"
            Write-Host "Download ist fehlgeschlagen. SciTE4AutoIt3.exe wurde nicht aktualisiert." -ForegroundColor "Red"
        }

        Write_LogEntry -Message "$($ProgramName) wurde aktualisiert." -Level "SUCCESS"
        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor "Green"
    } else {
        Write_LogEntry -Message "Kein Online Update verfügbar. $($ProgramName) ist aktuell." -Level "INFO"
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -ForegroundColor "DarkGray"
    }
} else {
    Write_LogEntry -Message "Konnte AutoIt Version oder Download-Link nicht extrahieren von der Seite." -Level "WARNING"
}

Write-Host ""
Write_LogEntry -Message "Prüfe erneut lokale Installationsdateien nach möglichen Änderungen." -Level "DEBUG"

$localAutoItFile = Get-ChildItem -Path "$InstallationFolder\autoit-v3-setup*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
$localSciTEFile = Get-ChildItem -Path "$InstallationFolder\SciTE4AutoIt3*.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1

Write_LogEntry -Message "Erneute Prüfung - lokale AutoIt Datei: $(if ($localAutoItFile) { $localAutoItFile.FullName } else { 'None' })" -Level "DEBUG"
Write_LogEntry -Message "Erneute Prüfung - lokale SciTE Datei: $(if ($localSciTEFile) { $localSciTEFile.FullName } else { 'None' })" -Level "DEBUG"

try {
    if ($localAutoItFile) {
        $localAutoItVersion = (Get-ItemProperty -Path $localAutoItFile.FullName).VersionInfo.FileVersion
        Write_LogEntry -Message "Ermittelte lokale AutoIt Version: $($localAutoItVersion)" -Level "DEBUG"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen AutoIt-Version: $($_)" -Level "ERROR"
}
try {
    if ($localSciTEFile) {
        $localSciTEVersion = (Get-ItemProperty -Path $localSciTEFile.FullName).VersionInfo.ProductVersion
        Write_LogEntry -Message "Ermittelte lokale SciTE Version: $($localSciTEVersion)" -Level "DEBUG"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Ermitteln der lokalen SciTE-Version: $($_)" -Level "ERROR"
}

#$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }

$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
Write_LogEntry -Message "Durchsuche Registry-Pfade: $($RegistryPaths -join ', ')" -Level "DEBUG"

$Path = foreach ($RegPath in $RegistryPaths) {
    if (Test-Path $RegPath) {
        Write_LogEntry -Message "Registry-Pfad existiert: $($RegPath)" -Level "DEBUG"
        Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$($ProgramName)*" }
    } else {
        Write_LogEntry -Message "Registry-Pfad nicht gefunden: $($RegPath)" -Level "DEBUG"
    }
}

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write_LogEntry -Message "$($ProgramName) ist installiert. Installierte Version: $($installedVersion); Installationsdatei Version: $($localAutoItVersion)" -Level "INFO"
    Write-Host "$ProgramName ist installiert." -ForegroundColor "Green"
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
    Write-Host "    Installationsdatei Version: $localAutoItVersion" -ForegroundColor "Cyan"

    if ([version]$installedVersion -lt [version]$localAutoItVersion) {
        Write_LogEntry -Message "Veraltete $($ProgramName) ist installiert. Update wird gestartet." -Level "INFO"
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "Magenta"
        $AutoItInstall = $true
    } elseif ([version]$installedVersion -eq [version]$localAutoItVersion) {
        Write_LogEntry -Message "Installierte Version ist aktuell." -Level "DEBUG"
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor "DarkGray"
        $AutoItInstall = $false
    } else {
        Write_LogEntry -Message "Installierte Version ($($installedVersion)) ist neuer als lokale Version ($($localAutoItVersion)). Kein Install nötig." -Level "WARNING"
        $AutoItInstall = $false
    }
} else {
    Write_LogEntry -Message "$($ProgramName) ist nicht in der Registry gefunden. Setze Install-Flag auf false." -Level "DEBUG"
    $AutoItInstall = $false
}
Write-Host ""

$Path = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like 'Scite*' }
Write_LogEntry -Message "Suche nach Scite Installation in Registry" -Level "DEBUG"

if ($null -ne $Path) {
    $installedVersion = $Path.DisplayVersion | Select-Object -First 1
    Write_LogEntry -Message "Scite ist installiert. Installierte Version: $($installedVersion); Installationsdatei Version: $($localSciTEVersion)" -Level "INFO"
    Write-Host "Scite ist installiert." -ForegroundColor "Green"
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor "Cyan"
    Write-Host "    Installationsdatei Version: $localSciTEVersion" -ForegroundColor "Cyan"

    if ([version]$installedVersion -lt [version]$localSciTEVersion) {
        Write_LogEntry -Message "Veraltete Scite ist installiert. Update wird gestartet." -Level "INFO"
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor "Magenta"
        $SciteInstall = $true
    } elseif ([version]$installedVersion -eq [version]$localSciTEVersion) {
        Write_LogEntry -Message "Installierte Scite Version ist aktuell." -Level "DEBUG"
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor "DarkGray"
        $SciteInstall = $false
    } else {
        Write_LogEntry -Message "Installierte Scite Version ($($installedVersion)) ist neuer als lokale Version ($($localSciTEVersion)). Kein Install nötig." -Level "WARNING"
        $SciteInstall = $false
    }
} else {
    Write_LogEntry -Message "Scite nicht in Registry gefunden." -Level "DEBUG"
    $SciteInstall = $false
}
Write-Host ""

if ($InstallationFlag) {
    Write_LogEntry -Message "Starte externes Installationsskript mit -InstallationFlag. Aufruf: $($PSHostPath) -File $($Serverip)\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1 -InstallationFlag" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1" `
        -InstallationFlag
    Write_LogEntry -Message "Externer Aufruf (InstallationFlag) beendet." -Level "DEBUG"
}

if ($AutoItInstall -eq $true) {
    Write_LogEntry -Message "Starte externes Installationsskript für AutoIt. Aufruf: $($PSHostPath) -File $($Serverip)\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1 -Autoit" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1" `
        -Autoit
    Write_LogEntry -Message "Externer AutoIt-Aufruf beendet." -Level "DEBUG"
}

if ($SciteInstall -eq $true) {
    Write_LogEntry -Message "Starte externes Installationsskript für SciTE. Aufruf: $($PSHostPath) -File $($Serverip)\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1 -Scite" -Level "INFO"
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Prog\InstallationScripts\Installation\AutoitInstallation.ps1" `
        -Scite
    Write_LogEntry -Message "Externer SciTE-Aufruf beendet." -Level "DEBUG"
}
Write-Host ""

Write_LogEntry -Message "Script endet normal." -Level "INFO"

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Script beendet: $($ProgramName) - $($ScriptType)" -Level "INFO"
    Finalize_LogSession | Out-Null
}
# === Ende Logger-Footer ===