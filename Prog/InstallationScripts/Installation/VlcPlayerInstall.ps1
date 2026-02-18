param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "VLC"
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

# DeployToolkit helpers
$dtPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Modules\DeployToolkit\DeployToolkit.psm1"
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
$configPath = Join-Path -Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Konfigurationspfad gesetzt: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei $($configPath) gefunden und importiert." -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"

$vlcSource = "$Serverip\Daten\Prog\vlc*.exe"
#$vlcDestination = "$env:USERPROFILE\Desktop"
$vlcShortcutPath = "$env:PUBLIC\Desktop\VLC media player.lnk"
$vlcStartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\VideoLAN"
Write_LogEntry -Message "VLC Quellpfad Pattern: $($vlcSource); ShortcutPath: $($vlcShortcutPath); StartMenuPath: $($vlcStartMenuPath)" -Level "DEBUG"

function SetFileAssociations($fileExtensions) {
    foreach ($extension in $fileExtensions) {
        Write_LogEntry -Message "Setze Dateizuordnung für Extension: $($extension) mit Programm: C:\Program Files\VideoLAN\VLC\vlc.exe" -Level "DEBUG"
        & $SetUserFTA --reg "C:\Program Files\VideoLAN\VLC\vlc.exe" $extension
    }
}

if (Test-Path $vlcSource) {
    Write-Host "VLC Player wird installiert" -foregroundcolor "magenta"
    Write_LogEntry -Message "VLC Installer Pattern vorhanden: $($vlcSource) (Test-Path true)" -Level "INFO"
    $vlcInstaller = Get-ChildItem -Path $vlcSource -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($vlcInstaller) {
        Write_LogEntry -Message "Gefundener VLC Installer: $($vlcInstaller.FullName)" -Level "DEBUG"
        Start-Process -FilePath $vlcInstaller.FullName -ArgumentList "/S" -Wait
        Write_LogEntry -Message "VLC Installer ausgeführt: $($vlcInstaller.FullName)" -Level "SUCCESS"
    } else {
        Write_LogEntry -Message "Kein konkreter VLC Installer unter Pattern $($vlcSource) gefunden." -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "VLC Installer Pattern nicht gefunden: $($vlcSource)" -Level "DEBUG"
}

if (Test-Path $vlcShortcutPath) {
    Write-Host "    Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $vlcShortcutPath -Force -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Public Desktop Shortcut entfernt: $($vlcShortcutPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Public Desktop Shortcut nicht gefunden: $($vlcShortcutPath)" -Level "DEBUG"
}

if (Test-Path $vlcStartMenuPath) {
    Write-Host "    Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $vlcStartMenuPath -Recurse -Force -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($vlcStartMenuPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Startmenüeintrag nicht gefunden: $($vlcStartMenuPath)" -Level "DEBUG"
}

$videoFileExtensions = @(
    ".3g2", ".3gp", ".3gp2", ".3gpp",
    ".asf", ".ASX", ".avi", ".M1V",
    ".m2t", ".m2ts", ".m4v", ".mkv",
    ".mov", ".MP2V", ".mp4", ".mp4v",
    ".mpa", ".mpe", ".mpeg", ".mpg",
    ".mpv2", ".mts", ".TS", ".TTS",
    ".wmv", ".wvx", ".m2t"
)

$musicFileExtensions = @(
    ".aac", ".adts", ".AIF", ".AIFC",
    ".AIFF", ".amr", ".AU", ".cda",
    ".flac", ".m3u", ".m4a", ".m4p",
    ".mid", ".mka", ".mp2", ".mp3",
    ".ra", ".ram", ".RMI", ".s3m",
    ".SND", ".voc", ".wav", ".wma",
    ".xm"
)

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag true: Setze Dateizuordnungen für VLC." -Level "INFO"
    if (Test-Path $SetUserFTA) {
        Write-Host "    VLC Video Dateizuordnung" -foregroundcolor "Yellow"
        Write_LogEntry -Message "SetUserFTA existiert: $($SetUserFTA). Starte Zuordnungen (Video)." -Level "DEBUG"
        SetFileAssociations $videoFileExtensions
        
        Write-Host "    VLC Music Dateizuordnung" -foregroundcolor "Yellow"
        Write_LogEntry -Message "Starte Zuordnungen (Music)." -Level "DEBUG"
        SetFileAssociations $musicFileExtensions
    } else {
        Write_LogEntry -Message "SetUserFTA nicht gefunden: $($SetUserFTA)" -Level "ERROR"
    }
    
    # === Start VLC, try to interact with first-run dialog if needed, then fallback to creating vlcrc ===

    # Pfad zur vlcrc
    $FilePath      = "$env:APPDATA\vlc\vlcrc"
    $vlcPath       = "C:\Program Files\VideoLAN\VLC\vlc.exe"

    # gewünschte Einstellungen (wie vorher)
    $desired = @{
        'waveout-volume'   = 'waveout-volume=0.500000'
        'mmdevice-volume'  = 'mmdevice-volume=0.500000'
        'directx-volume'   = 'directx-volume=0.500000'
        'volume-save'      = 'volume-save=0'
    }

    # helper: ensure config dir exists
    $vlcConfigDir = Split-Path -Path $FilePath -Parent
    if (-not (Test-Path -Path $vlcConfigDir)) {
        try { New-Item -Path $vlcConfigDir -ItemType Directory -Force | Out-Null } catch { Write_LogEntry -Message "Konnte Verzeichnis für vlcrc nicht anlegen: $($vlcConfigDir) - $_" -Level "WARNING" }
    }

    # Start VLC if available
    $timeoutTotal  = 60      # total seconds to wait for vlcrc to appear
    $interactiveWait = 1     # seconds to wait before trying GUI interaction
    $pollInterval  = 1
    $elapsed       = 0

    if (Test-Path $vlcPath) {
        Write_LogEntry -Message "Starte VLC zum Erzeugen der vlcrc: $($vlcPath)" -Level "INFO"
        Start-Process -FilePath $vlcPath
        Write_LogEntry -Message "VLC Prozess gestartet (Start-Process aufgerufen): $($vlcPath)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "VLC Pfad nicht gefunden: $($vlcPath). Überspringe Start." -Level "ERROR"
    }

    # Passive wait for small time before GUI interaction attempt
    while (($elapsed -lt $interactiveWait) -and (-not (Test-Path -Path $FilePath))) {
        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
    }

    if (Test-Path -Path $FilePath) {
        Write_LogEntry -Message "vlcrc früh gefunden (nach $elapsed s)." -Level "DEBUG"
    } else {
        Write_LogEntry -Message "vlcrc noch nicht vorhanden nach $elapsed s. Versuche GUI-Interaktion (wenn Fenster vorhanden)..." -Level "DEBUG"

        # Wait for a process with a main window to appear
        $waitForWindow = 10
        $winElapsed     = 0
        $proc = $null
        do {
            Start-Sleep -Milliseconds 300
            $proc = Get-Process -Name vlc -ErrorAction SilentlyContinue
            $winElapsed += 0.3
        } while (($proc -eq $null -or $proc.MainWindowHandle -eq 0) -and ($winElapsed -lt $waitForWindow))

        if ($proc -and $proc.MainWindowHandle -ne 0) {
            Write_LogEntry -Message "VLC Fenster gefunden; PID: $($proc.Id); Handle: $($proc.MainWindowHandle)" -Level "DEBUG"

            # Bring window to foreground
            Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@ -ErrorAction SilentlyContinue

            try {
                [Win32]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
                Start-Sleep -Milliseconds 400

                # load SendKeys assembly and send keys to accept dialogs then close
                Add-Type -AssemblyName System.Windows.Forms
                # Press ENTER to confirm dialog (e.g. first-run)
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                Start-Sleep -Milliseconds 400
                # Then Alt+F4 to close VLC
                [System.Windows.Forms.SendKeys]::SendWait("%{F4}")
                Write_LogEntry -Message "Sende Tasten an VLC (ENTER, Alt+F4) zur Bestätigung/Schließen." -Level "INFO"

                # Give a short moment for vlcrc to be created
                Start-Sleep -Seconds 2
            } catch {
                Write_LogEntry -Message "Fehler während GUI-Interaktion: $($_)" -Level "WARNING"
            }
        } else {
            Write_LogEntry -Message "Kein sichtbares VLC-Fenster gefunden innerhalb $waitForWindow s; überspringe GUI-Interaktion." -Level "DEBUG"
        }

        # Continue polling for the remainder of the total timeout
        while (($elapsed -lt $timeoutTotal) -and (-not (Test-Path -Path $FilePath))) {
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval
        }

        if (Test-Path -Path $FilePath) {
            Write_LogEntry -Message "vlcrc gefunden nach GUI-Interaktion / Polling (gesamt $elapsed s)." -Level "DEBUG"
        } else {
            Write_LogEntry -Message "vlcrc nicht gefunden nach $elapsed s Gesamtwartezeit. Erzeuge Fallback." -Level "WARNING"

            # attempt to stop VLC if still running
            try {
                $procKill = Get-Process -Name vlc -ErrorAction SilentlyContinue
                if ($procKill) {
                    Write_LogEntry -Message "VLC noch aktiv; versuche Stop-Process PID: $($procKill.Id)" -Level "DEBUG"
                    $procKill | Stop-Process -Force -ErrorAction SilentlyContinue
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Beenden des VLC Prozesses: $($_)" -Level "DEBUG"
            }

            # Create deterministic fallback vlcrc with the desired keys
            $defaultLines = @()
            $defaultLines += "# vlcrc automatisch erzeugt vom Installationsskript - Fallback"
            $defaultLines += "# Datum: $(Get-Date -Format 'u')"
            foreach ($k in $desired.Keys) { $defaultLines += $desired[$k] }

            try {
                Set-Content -LiteralPath $FilePath -Value $defaultLines -Encoding UTF8 -Force
                Write_LogEntry -Message "Fallback vlcrc geschrieben: $FilePath" -Level "INFO"
            } catch {
                Write_LogEntry -Message "Fehler beim Erstellen der Fallback vlcrc: $($_)" -Level "ERROR"
            }
        }
    }

    # Now edit (or add) desired keys into vlcrc safely
    if (Test-Path -Path $FilePath) {
        try {
            [string[]]$lines = Get-Content -LiteralPath $FilePath -Encoding UTF8
            Write_LogEntry -Message "vlcrc Datei eingelesen: $($FilePath); Zeilen: $($lines.Count)" -Level "DEBUG"

            for ($i = 0; $i -lt $lines.Count; $i++) {
                $trimmed = $lines[$i].TrimStart()
                foreach ($key in $desired.Keys) {
                    if ($trimmed -match "^(#\s*)?$([regex]::Escape($key))\s*=") {
                        $lines[$i] = $desired[$key]
                        break
                    }
                }
            }

            # If keys not present, append them
            foreach ($key in $desired.Keys) {
                $exists = $false
                foreach ($l in $lines) { if ($l -match "^\s*$([regex]::Escape($key))\s*=") { $exists = $true; break } }
                if (-not $exists) { $lines += $desired[$key]; Write_LogEntry -Message "Key hinzugefügt in vlcrc: $($key)" -Level "DEBUG" }
            }

            Set-Content -LiteralPath $FilePath -Value $lines -Encoding UTF8 -Force
            Write_LogEntry -Message "vlcrc Datei überschrieben mit neuen Einstellungen: $($FilePath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Lesen/Schreiben der vlcrc: $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "vlcrc nicht vorhanden und konnte nicht erzeugt. Überspringe Anpassung." -Level "ERROR"
    }

    # Abschließend: falls noch ein VLC Prozess offen ist (weil wir wollten, dass vlcrc generiert wird), schließe ihn jetzt
    try {
        $procNow = Get-Process -Name vlc -ErrorAction SilentlyContinue
        if ($procNow) {
            $procNow | Stop-Process -Force -ErrorAction SilentlyContinue
            Write_LogEntry -Message "VLC Prozess beendet nach Konfiguration (wenn vorhanden)." -Level "DEBUG"
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Beenden des VLC Prozesses am Ende: $($_)" -Level "DEBUG"
    }

    Write-Host "    VLC volume Einstellungen sind geändert." -foregroundcolor "Green"
    Write_LogEntry -Message "VLC volume Einstellungen gesetzt: $($FilePath)" -Level "INFO"
}

#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\VLC_audio_entfernen.reg" -Wait
#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\VLC_other_entfernen.reg" -Wait
#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\VLC_video_entfernen.reg" -Wait

Write_LogEntry -Message "Importiere Registry-Kontextmenüs via RegistryImport Skript." -Level "INFO"
Write_LogEntry -Message "Aufruf: RegistryImport.ps1 für VLC_audio_entfernen.reg" -Level "DEBUG"
& $PSHostPath `
    -NoLogo -NoProfile -ExecutionPolicy Bypass `
    -File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
    -Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\VLC_audio_entfernen.reg"
Write_LogEntry -Message "Beendet: RegistryImport.ps1 für VLC_audio_entfernen.reg" -Level "DEBUG"

Write_LogEntry -Message "Aufruf: RegistryImport.ps1 für VLC_other_entfernen.reg" -Level "DEBUG"
& $PSHostPath `
    -NoLogo -NoProfile -ExecutionPolicy Bypass `
    -File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
    -Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\VLC_other_entfernen.reg"
Write_LogEntry -Message "Beendet: RegistryImport.ps1 für VLC_other_entfernen.reg" -Level "DEBUG"

Write_LogEntry -Message "Aufruf: RegistryImport.ps1 für VLC_video_entfernen.reg" -Level "DEBUG"
& $PSHostPath `
    -NoLogo -NoProfile -ExecutionPolicy Bypass `
    -File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
    -Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\VLC_video_entfernen.reg"
Write_LogEntry -Message "Beendet: RegistryImport.ps1 für VLC_video_entfernen.reg" -Level "DEBUG"

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
