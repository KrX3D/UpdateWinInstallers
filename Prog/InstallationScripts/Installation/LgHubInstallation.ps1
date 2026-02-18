param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Logitech G HUB"
$ScriptType  = "Install"

# === Logger-Header: automatisch eingefügt ===
$parentPath  = Split-Path -Path $PSScriptRoot -Parent
$modulePath  = Join-Path -Path $parentPath -ChildPath 'Modules\Logger\Logger.psm1'

if (Test-Path $modulePath) {
    Import-Module -Name $modulePath -Force -ErrorAction Stop

    if (-not (Get-Variable -Name logRoot -Scope Script -ErrorAction SilentlyContinue)) {
        $logRoot = Join-Path -Path $parentPath -ChildPath 'Log'
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
$configPath = Join-Path -Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Konfigurationspfad gesetzt: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei $($configPath) gefunden und importiert." -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Script beendet wegen fehlender Konfigurationsdatei: $($configPath)" -Level "ERROR"
    exit
}

$LgHubFolder = "$Serverip\Daten\Treiber\LgHub"
Write_LogEntry -Message "LgHubFolder gesetzt: $($LgHubFolder)" -Level "DEBUG"

#Uninstall
if ($InstallationFlag -eq $false) {
    $lgHubExePath = Join-Path $env:ProgramFiles 'LGHUB\lghub_updater.exe'
    Write_LogEntry -Message "Uninstall-Pfad geprüft: $($lgHubExePath)" -Level "DEBUG"

    # Check if LG Hub is currently running
    function IsLGHubRunning {
        return Get-Process -Name 'lghub_updater' -ErrorAction SilentlyContinue
    }

    # Check if LG Hub is installed
    if (Test-Path $lgHubExePath) {
        # Uninstall LG Hub
        Write_LogEntry -Message "Vorhandene Logitech G Hub Version gefunden: $($lgHubExePath). Starte Deinstallation." -Level "INFO"
        Write-Host "Vorhandene Logitech G Hub Version wird deinstalliert." -foregroundcolor "magenta"

        try {
            Start-Process -FilePath $lgHubExePath -ArgumentList "--uninstall" -Wait
            Write_LogEntry -Message "Deinstallationsprozess gestartet für: $($lgHubExePath)" -Level "INFO"
        } catch {
            Write_LogEntry -Message "Fehler beim Starten der Deinstallation $($lgHubExePath): $($_)" -Level "ERROR"
        }

        # Wait until LG Hub is uninstalled
        while (IsLGHubRunning) {
            Write_LogEntry -Message "Warte bis LG Hub deinstalliert ist..." -Level "DEBUG"
            Write-Host "Warte bis LG Hub deinstalliert ist..." -foregroundcolor "Yellow"
            Start-Sleep -Seconds 5
        }

        Write_LogEntry -Message "LG Hub Deinstallation abgeschlossen (keine Prozesse mehr)." -Level "SUCCESS"
        Write-Host "LG Hub wurde deinstalliert." -foregroundcolor "Cyan"
        Write-Host ""

        start-Sleep -Milliseconds 1000

        $folderPath = "C:\ProgramData\LGHUB"
        Write_LogEntry -Message "Prüfe verbleibende Ordner: $($folderPath)" -Level "DEBUG"

        if (Test-Path -Path $folderPath) {
            try {
                Remove-Item -Path $folderPath -Force -Recurse
                Write_LogEntry -Message "Ordner entfernt: $($folderPath)" -Level "SUCCESS"
                start-Sleep -Milliseconds 1000
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen des Ordners $($folderPath): $($_)" -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Kein Ordner gefunden zum Entfernen: $($folderPath)" -Level "DEBUG"
        }
    } else {
        Write_LogEntry -Message "Keine installierte Logitech G Hub Version gefunden unter: $($lgHubExePath)" -Level "INFO"
    }
}

#Install
Write_LogEntry -Message "Starte Installations-Block: Logitech G Hub" -Level "INFO"
Write-Host "Logitech G Hub wird installiert" -foregroundcolor "magenta"

$lgHubInstallerPath = Get-ChildItem -Path "$Serverip\Daten\Treiber\LgHub\lghub_installer_*.exe" | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue
Write_LogEntry -Message "Gefundene Installerquelle: $($lgHubInstallerPath)" -Level "DEBUG"

if ($lgHubInstallerPath) {
    try {
        Copy-Item -Path $lgHubInstallerPath -Destination "$env:TEMP\" -Force
        Write_LogEntry -Message "Installer kopiert nach Temp: $($env:TEMP) (Quelle: $($lgHubInstallerPath))" -Level "INFO"
    } catch {
        Write_LogEntry -Message "Fehler beim Kopieren des Installers $($lgHubInstallerPath) nach $($env:TEMP): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Installer unter $($LgHubFolder) gefunden (Pattern lghub_installer_*.exe)." -Level "ERROR"
}

$installerPath = Get-ChildItem -Path "$env:TEMP\lghub_installer_*.exe" | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue
Write_LogEntry -Message "Installer im Temp geprüft: $($installerPath)" -Level "DEBUG"

if ($installerPath) {
    try {
        Start-Process -FilePath $installerPath -ArgumentList "--silent"
        Write_LogEntry -Message "Start-Process ausgeführt für Installer: $($installerPath) mit Argument '--silent' (asynchron)" -Level "INFO"
    } catch {
        Write_LogEntry -Message "Fehler beim Starten des Installers $($installerPath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Installer im Temp gefunden, überspringe Start." -Level "WARNING"
}

# Silent installer autostart app. Killing this behaviour using the code below
#https://community.chocolatey.org/packages/advanced-port-scanner#files
$t = 0
Write_LogEntry -Message "Warte-Schleife startet, Prüfung auf Prozess 'lghub' (Timeout 90s)..." -Level "DEBUG"
DO
{
    start-Sleep -Milliseconds 1000
    $t++ #increase iteration 
} Until ($null -ne ($p=Get-Process -Name lghub -ErrorAction SilentlyContinue) -or ($t -gt 90)) #wait until process is found or timeout reached

if($p) { #if process is found
    try {
        $procCount = @($p).Count
        Write_LogEntry -Message "Gefundene lghub Prozessanzahl nach Installerstart: $($procCount)" -Level "DEBUG"
        $p | Stop-Process  -Force
        Write_LogEntry -Message "Beende Logitech G Hub Prozess( e ). Prozessanzahl: $($procCount)" -Level "INFO"
        Write-Host "Beende Logitech G Hub Prozess"
    } catch {
        Write_LogEntry -Message "Fehler beim Beenden des/neuer Prozesse: $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Timeout erreicht: Kein lghub Prozess gefunden innerhalb von $($t) Sekunden." -Level "WARNING"
    Write-Host "Timeout für Logitech G Hub Prozess" #no process found but timeout reached
}

if (Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Logi") {
    Write_LogEntry -Message "Startmenüeintrag 'Logi' gefunden, wird entfernt." -Level "INFO"
    Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    try {
        Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Logi" -Recurse -Force
        Write_LogEntry -Message "Startmenüeintrag entfernt: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Logi" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Entfernen des Startmenüeintrags 'Logi': $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag 'Logi' gefunden." -Level "DEBUG"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag gesetzt: Zusätzliche Schritte werden ausgeführt (LG Prozesse stoppen, Konfiguration wiederherstellen)." -Level "INFO"
    Write-Host "	LG Prozesse werden beendet." -foregroundcolor "Yellow"
    $lgHubProcesses = Get-Process | Where-Object { $_.ProcessName -like "*lghub*" }
    $procCountDuringRestore = @($lgHubProcesses).Count
    Write_LogEntry -Message "Gefundene lghub Prozesse vor Wiederherstellung: $($procCountDuringRestore)" -Level "DEBUG"

    if ($lgHubProcesses) {
        $lgHubProcesses | ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                Write_LogEntry -Message "Beende Prozess Id: $($_.Id) Name: $($_.ProcessName)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Beenden von Prozess Id $($_.Id): $($_)" -Level "ERROR"
            }
        }
    } else {
        Write_LogEntry -Message "Keine lghub Prozesse zum Beenden gefunden." -Level "DEBUG"
    }

    $sourceLocal = "$LgHubFolder\Local\LGHUB"
    $destinationLocal = "$env:USERPROFILE\AppData\Local\LGHub"
    Write_LogEntry -Message "Lokale Konfig-Pfade: Quelle=$($sourceLocal) Ziel=$($destinationLocal)" -Level "DEBUG"

    if (Test-Path $sourceLocal) {
        Write_LogEntry -Message "Quelle für lokale Konfiguration vorhanden: $($sourceLocal). Starte Kopieren." -Level "INFO"
        if (!(Test-Path $destinationLocal)) {
            New-Item -ItemType Directory -Path $destinationLocal -Force | Out-Null
            Write_LogEntry -Message "Zielverzeichnis erstellt: $($destinationLocal)" -Level "DEBUG"
        }
        try {
            Get-ChildItem $sourceLocal | Copy-Item -Destination $destinationLocal -Recurse -Force
            Write_LogEntry -Message "Lokale Konfiguration kopiert von $($sourceLocal) nach $($destinationLocal)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Kopieren lokaler Konfiguration von $($sourceLocal) nach $($destinationLocal): $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Keine lokale Konfiguration gefunden unter: $($sourceLocal)" -Level "WARNING"
    }

    $sourceRoaming = "$LgHubFolder\Roaming\LGHUB"
    $destinationRoaming = "$env:USERPROFILE\AppData\Roaming\LGHub"
    Write_LogEntry -Message "Roaming Konfig-Pfade: Quelle=$($sourceRoaming) Ziel=$($destinationRoaming)" -Level "DEBUG"

    if (Test-Path $sourceRoaming) {
        Write_LogEntry -Message "Quelle für roaming Konfiguration vorhanden: $($sourceRoaming). Starte Kopieren." -Level "INFO"
        if (!(Test-Path $destinationRoaming)) {
            New-Item -ItemType Directory -Path $destinationRoaming -Force | Out-Null
            Write_LogEntry -Message "Zielverzeichnis erstellt: $($destinationRoaming)" -Level "DEBUG"
        }
        try {
            Get-ChildItem $sourceRoaming | Copy-Item -Destination $destinationRoaming -Recurse -Force
            Write_LogEntry -Message "Roaming Konfiguration kopiert von $($sourceRoaming) nach $($destinationRoaming)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Kopieren roaming Konfiguration von $($sourceRoaming) nach $($destinationRoaming): $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Keine roaming Konfiguration gefunden unter: $($sourceRoaming)" -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "InstallationFlag nicht gesetzt: Keine Wiederherstellung der LG Konfiguration." -Level "DEBUG"
}

start-Sleep -Milliseconds 3000
Write_LogEntry -Message "Kurze Pause nach Installationsschritten (3000ms) abgeschlossen." -Level "DEBUG"

# Remove the Installer from Temp if it exists
if (Test-Path $installerPath) {
    try {
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        Write_LogEntry -Message "Installer aus Temp entfernt: $($installerPath)" -Level "INFO"
    } catch {
        Write_LogEntry -Message "Fehler beim Entfernen des Installers $($installerPath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Installer in Temp vorhanden zum Entfernen: $($installerPath)" -Level "DEBUG"
}

# Ensure C:\ProgramData\LGHUB\periodic_check.json has downloadAutomatically = false
try {
    $periodicFile = 'C:\ProgramData\LGHUB\periodic_check.json'
    Write_LogEntry -Message "Prüfe Datei: $periodicFile" -Level "DEBUG"

    if (Test-Path -Path $periodicFile) {
        try {
            $raw = Get-Content -Path $periodicFile -Raw -ErrorAction Stop
            $json = $raw | ConvertFrom-Json -ErrorAction Stop

            if ($null -eq $json.downloadAutomatically -or $json.downloadAutomatically -ne $false) {
                $oldVal = $json.downloadAutomatically
                $json.downloadAutomatically = $false
                $json | ConvertTo-Json -Depth 10 | Set-Content -Path $periodicFile -Encoding UTF8
                Write_LogEntry -Message "periodic_check.json aktualisiert: downloadAutomatically $oldVal -> $($json.downloadAutomatically)" -Level "SUCCESS"
                Write-Host "Updated $periodicFile : downloadAutomatically set to false" -ForegroundColor Cyan
            } else {
                Write_LogEntry -Message "periodic_check.json bereits korrekt (downloadAutomatically = false)." -Level "DEBUG"
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Einlesen/Verarbeiten von $periodicFile : $($_.Exception.Message)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "periodic_check.json nicht gefunden unter $periodicFile . Nichts zu ändern." -Level "DEBUG"
    }
} catch {
    Write_LogEntry -Message "Unerwarteter Fehler beim Prüfen/Ändern von periodic_check.json: $($_.Exception.Message)" -Level "ERROR"
}

# --- Run Set-GHubSetting helper (no params) and ensure periodic_check.json has downloadAutomatically=false ---
try {
    $setScriptPath = Join-Path -Path $Serverip -ChildPath 'Daten\Customize_Windows\Scripte\LgHub\Set-GHubSetting.ps1'
    Write_LogEntry -Message "Prüfe Vorhandensein von Set-GHubSetting: $setScriptPath" -Level "DEBUG"

    if (Test-Path -Path $setScriptPath) {
        Write_LogEntry -Message "Starte Set-GHubSetting.ps1 und warte auf Abschluss..." -Level "INFO"
        try {
            $proc = Start-Process -FilePath 'powershell.exe' `
                                  -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$setScriptPath) `
                                  -WindowStyle Hidden -PassThru -Wait -ErrorAction Stop
            $exitCode = $proc.ExitCode
            Write_LogEntry -Message "Set-GHubSetting.ps1 beendet mit ExitCode: $exitCode" -Level "INFO"
        } catch {
            Write_LogEntry -Message "Fehler beim Starten/ausführen von Set-GHubSetting.ps1: $($_.Exception.Message)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Set-GHubSetting.ps1 nicht gefunden unter: $setScriptPath" -Level "WARNING"
    }
} catch {
    Write_LogEntry -Message "Unerwarteter Fehler beim Ausführen von Set-GHubSetting.ps1: $($_.Exception.Message)" -Level "ERROR"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
