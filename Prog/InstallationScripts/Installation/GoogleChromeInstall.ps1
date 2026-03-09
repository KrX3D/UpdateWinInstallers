param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Google Chrome"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt: $($ProgramName); ScriptType gesetzt: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

#Bei Update wird ohne deinstallation die neue Version richtig installiert

Write-Host "Google Chrome wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Starte Chrome-Installationsabschnitt" -Level "INFO"

# Find MSI — choose the newest file if multiple exist
try {
    $chromeMsiCandidates = Get-ChildItem -Path "$Serverip\Daten\Prog\*Chrome*.msi" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if ($chromeMsiCandidates -and $chromeMsiCandidates.Count -gt 0) {
        $chromeMsi = $chromeMsiCandidates[0].FullName
        Write_LogEntry -Message "Gefundene Chrome MSI (aus Kandidaten): $($chromeMsi)" -Level "DEBUG"
    } else {
        $chromeMsi = $null
        Write_LogEntry -Message "Keine Chrome MSI unter $($Serverip)\Daten\Prog gefunden." -Level "ERROR"
    }
} catch {
    $chromeMsi = $null
    Write_LogEntry -Message "Fehler beim Suchen der Chrome MSI: $($_)" -Level "ERROR"
}

if (-not $chromeMsi) {
    Write-Host "Keine Chrome MSI gefunden, Script abgebrochen." -ForegroundColor Red
    exit 1
}

# Install via msiexec (quiet)
Write_LogEntry -Message "Starte msiexec zur Installation von: $($chromeMsi)" -Level "INFO"
$msiArgs = "/i `"$chromeMsi`" /quiet /norestart"
$proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -ErrorAction SilentlyContinue
if ($proc -and $proc.ExitCode -ne $null) {
    Write_LogEntry -Message "msiexec beendet für: $($chromeMsi) ExitCode: $($proc.ExitCode)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "msiexec wurde gestartet (kein ExitCode verfügbar) oder Fehler beim Start." -Level "WARNING"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag = true: Beginne Chrome-Konfiguration" -Level "INFO"
    Write-Host "Google Chrome wird eingestellt." -foregroundcolor "Cyan"

    # Start Chrome (to create shortcuts/profile files if needed)
    Write_LogEntry -Message "Starte Chrome (für Shortcut/Profil-Konfiguration)" -Level "DEBUG"
    $chromeExePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    if (-not (Test-Path -Path $chromeExePath)) {
        # On some systems Chrome may be in Program Files (x86)
        $chromeExePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    }

    if (Test-Path -Path $chromeExePath) {
        Start-Process -FilePath $chromeExePath -ErrorAction SilentlyContinue
        Write_LogEntry -Message "Chrome Start-Aufruf durchgeführt: $($chromeExePath)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Chrome executable nicht gefunden: $($chromeExePath). Überspringe Start." -Level "WARNING"
    }

    # small wait to let Chrome create initial files/shortcuts
    Start-Sleep -Seconds 5
    Write_LogEntry -Message "Warte 5 Sekunden nach Chrome-Start" -Level "DEBUG"

    # If Chrome is running, stop it so we can copy profile files safely
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Get-Process chrome ergab: $($($chromeProcess).Count) Prozesse" -Level "DEBUG"

    if ($chromeProcess) {
        Write_LogEntry -Message "Vorhandene Chrome-Prozesse gefunden; Stop-Process wird ausgeführt." -Level "INFO"
        try {
            Stop-Process -Name "chrome" -Force -ErrorAction Stop
            Write_LogEntry -Message "Stop-Process für chrome ausgeführt." -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Stop-Process chrome: $($_)" -Level "WARNING"
        }
    } else {
        Write_LogEntry -Message "Keine Chrome-Prozesse gefunden vor Stop-Process." -Level "DEBUG"
    }

    # Retrieve machine name via external script (robust handling)
    $scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
    Write_LogEntry -Message "Hole PCName via externes Script: $($scriptPath)" -Level "INFO"

    try {
		#$PCName = & $scriptPath
		#$PCName = & "powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
		$PCName = & $PSHostPath `
			-NoLogo -NoProfile -ExecutionPolicy Bypass `
			-File $scriptPath `
			-Verbose:$false
        Write_LogEntry -Message "Externes Script $($scriptPath) ausgeführt; PCName=$($PCName)" -Level "INFO"
	} catch {
		Write_LogEntry -Message "Failed to load script $($scriptPath). Reason: $($($_))" -Level "ERROR"
		Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
		Pause
		Exit
	}

    # Stop Explorer to ensure public Desktop/StartMenu modifications are clean
    Write_LogEntry -Message "Stoppe Explorer vor Weiterarbeit (vorher evtl. Wiederherstellung)" -Level "INFO"
    try {
        Stop-Process -Name explorer -Force -ErrorAction Stop
        Write_LogEntry -Message "Explorer gestoppt." -Level "DEBUG"
    } catch {
        Write_LogEntry -Message "Fehler beim Stoppen von Explorer (Ignoriere): $($_)" -Level "WARNING"
    }
    Start-Sleep -Seconds 5
    Write_LogEntry -Message "Warte 5 Sekunden nach Stop-Process explorer" -Level "DEBUG"
	
    # Set file associations via SFTA (if available)
    $SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
    Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"
#https://github.com/DanysysTeam/PS-SFTA/blob/master/SFTA.ps1

	#$fileTypes = @(".html")
	$fileTypes = @(
		".htm",
		".html",
		".shtml",
		".svg",
		".webp",
		".xht",
		".xhtml"
	)

	# Also register web protocols
	$protocols = @(
		"http",
		"https",
		"ftp"
	)
	
	if (Test-Path $SetUserFTA -PathType Leaf) {
		Write_LogEntry -Message "SFTA Tool gefunden: $($SetUserFTA). Setze Dateizuordnungen." -Level "INFO"
		Write-Host "Chrome Dateizuordnung" -foregroundcolor "Yellow"
		
		$chromeExe = "C:\Program Files\Google\Chrome\Application\chrome.exe"
		
		# Set file types
		foreach ($type in $fileTypes | Sort-Object) {
			Write_LogEntry -Message "Rufe SFTA auf für Typ: $($type)" -Level "DEBUG"
			try {
				& $SetUserFTA --reg $chromeExe $type
				Write_LogEntry -Message "SFTA-Aufruf beendet für Typ: $($type)" -Level "SUCCESS"
			} catch {
				Write_LogEntry -Message "SFTA-Aufruf für $($type) fehlgeschlagen: $($_)" -Level "WARNING"
			}
		}

		# Set protocols
		foreach ($protocol in $protocols | Sort-Object) {
			Write_LogEntry -Message "Rufe SFTA auf für Protokoll: $($protocol)" -Level "DEBUG"
			try {
				& $SetUserFTA --reg $chromeExe $protocol
				Write_LogEntry -Message "SFTA-Aufruf beendet für Protokoll: $($protocol)" -Level "SUCCESS"
			} catch {
				Write_LogEntry -Message "SFTA-Aufruf für $($protocol) fehlgeschlagen: $($_)" -Level "WARNING"
			}
        }
    } else {
        Write_LogEntry -Message "SFTA Tool nicht gefunden: $($SetUserFTA)" -Level "WARNING"
    }

    # Profile restore: robust copy (robocopy preferred, fallback to Copy-Item)
    if ($PCName -ne "Unknown") {
        Write_LogEntry -Message "PCName erkannt: $($PCName). Versuche Chrome-Profil-Wiederherstellung." -Level "INFO"
        $ChromeProfileFolder = Join-Path -Path "$Serverip\Daten\Prog\GoogleChrome" -ChildPath $PCName
        $ChromeProfileFolder = Join-Path -Path $ChromeProfileFolder -ChildPath "Default"
        $destination = "$env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default"

        Write_LogEntry -Message "Quell-Profil: $($ChromeProfileFolder); Ziel: $($destination)" -Level "DEBUG"

        if (Test-Path -Path $ChromeProfileFolder) {

            # Ensure destination exists
            try {
                if (-not (Test-Path -Path $destination)) {
                    New-Item -ItemType Directory -Path $destination -Force | Out-Null
                    Write_LogEntry -Message "Zielverzeichnis erstellt: $($destination)" -Level "DEBUG"
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Erstellen des Zielverzeichnisses: $($_)" -Level "ERROR"
            }

            # Use Robocopy for robust copying (preserves attributes and can handle locked files better)
            $robocopyExe = (Get-Command robocopy.exe -ErrorAction SilentlyContinue).Source
            if ($robocopyExe) {
                # Use trailing backslash on source to copy contents into destination
                $srcForRobo = $ChromeProfileFolder.TrimEnd('\') + '\'
                $destForRobo = $destination.TrimEnd('\') + '\'

                # Build args (UNQUOTED). We'll run robocopy via & so PowerShell handles argument passing correctly.
                $roboParams = @($srcForRobo, $destForRobo, '/E', '/COPY:DAT', '/R:3', '/W:5', '/NFL', '/NDL', '/NJH', '/NJS')

                Write_LogEntry -Message "Robocopy wird verwendet: $($robocopyExe) $($roboParams -join ' ')" -Level "DEBUG"

                try {
                    # Run robocopy and capture output
                    $roboOutput = & $robocopyExe @roboParams 2>&1
                    $roboExit = $LASTEXITCODE

                    # Log robocopy output (first 200 lines max to avoid huge logs)
                    if ($roboOutput) {
                        $roboOutput[0..([Math]::Min($roboOutput.Count-1,199))] | ForEach-Object { Write_LogEntry -Message "Robocopy: $_" -Level "DEBUG" }
                    } else {
                        Write_LogEntry -Message "Robocopy lieferte keine Ausgabe." -Level "DEBUG"
                    }

                    Write_LogEntry -Message "Robocopy ExitCode: $($roboExit)" -Level "DEBUG"

                    if ($roboExit -le 7) {
                        Write_LogEntry -Message "Robocopy erfolgreich: $($ChromeProfileFolder) -> $($destination)" -Level "SUCCESS"
                    } else {
                        Write_LogEntry -Message "Robocopy meldete Fehler (ExitCode $roboExit). Versuche Fallback Copy-Item." -Level "WARNING"
                        throw "RobocopyExitCode:$roboExit"
                    }
                } catch {
                    Write_LogEntry -Message "Robocopy fehlgeschlagen oder ExitCode ungünstig: $($_). Versuche Copy-Item Fallback." -Level "WARNING"
                    # Fallback: Copy-Item approach
                    try {
                        Get-ChildItem -Path $ChromeProfileFolder -Force -Recurse | ForEach-Object {
                            $targetPath = Join-Path -Path $destination -ChildPath ($_.FullName.Substring($ChromeProfileFolder.Length).TrimStart('\'))
                            $targetDir = Split-Path -Path $targetPath -Parent
                            if (-not (Test-Path -Path $targetDir)) { New-Item -Path $targetDir -ItemType Directory -Force | Out-Null }
                            Copy-Item -LiteralPath $_.FullName -Destination $targetPath -Force -ErrorAction Stop
                        }
                        Write_LogEntry -Message "Fallback Copy-Item erfolgreich: $($ChromeProfileFolder) -> $($destination)" -Level "SUCCESS"
                    } catch {
                        Write_LogEntry -Message "Fallback Copy-Item fehlgeschlagen: $($_)" -Level "ERROR"
                    }
                }
            } else {
                # No robocopy: fallback to Copy-Item with retries
                Write_LogEntry -Message "Robocopy nicht verfügbar, benutze Copy-Item mit Retries." -Level "DEBUG"
                $maxAttempts = 3
                $attempt = 0
                $copied = $false
                while (($attempt -lt $maxAttempts) -and (-not $copied)) {
                    $attempt++
                    try {
                        Get-ChildItem -Path $ChromeProfileFolder -Force -Recurse | ForEach-Object {
                            $targetPath = Join-Path -Path $destination -ChildPath ($_.FullName.Substring($ChromeProfileFolder.Length).TrimStart('\'))
                            $targetDir = Split-Path -Path $targetPath -Parent
                            if (-not (Test-Path -Path $targetDir)) { New-Item -Path $targetDir -ItemType Directory -Force | Out-Null }
                            Copy-Item -LiteralPath $_.FullName -Destination $targetPath -Force -ErrorAction Stop
                        }
                        $copied = $true
                        Write_LogEntry -Message "Copy-Item erfolgreich nach Versuch $attempt." -Level "SUCCESS"
                    } catch {
                        Write_LogEntry -Message "Copy-Item Versuch $attempt fehlgeschlagen: $($_). Warte 3s und retry." -Level "WARNING"
                        Start-Sleep -Seconds 3
                    }
                }
                if (-not $copied) {
                    Write_LogEntry -Message "Alle Copy-Item Versuche fehlgeschlagen." -Level "ERROR"
                }
            }
        } else {
            Write_LogEntry -Message "Quell-Profil nicht gefunden: $($ChromeProfileFolder). Überspringe Kopie." -Level "INFO"
        }
    } else {
        Write_LogEntry -Message "PCName ist 'Unknown', überspringe Profil-Wiederherstellung" -Level "WARNING"
    }

    # restart explorer so desktop/startmenu is back
    #try {
        #Start-Process -FilePath "explorer.exe" -ErrorAction SilentlyContinue
        #Write_LogEntry -Message "Explorer neu gestartet." -Level "DEBUG"
    #} catch {
        #Write_LogEntry -Message "Fehler beim Starten von Explorer: $($_)" -Level "WARNING"
    #}
	
	# Stop Explorer
	Stop-Process -Name "explorer" -Force -ErrorAction Stop
	Write_LogEntry -Message "Explorer gestoppt." -Level "DEBUG"

	# Wait for Windows to auto-restart explorer (it will)
	Start-Sleep -Seconds 3

	# Now close any Explorer windows that opened (but keep shell running)
	try {
		$shell = New-Object -ComObject Shell.Application
		$windows = $shell.Windows()
		
		foreach ($window in $windows) {
			try {
				# Close any folder windows (This PC, etc.)
				$window.Quit()
			} catch {
				# Ignore
			}
		}
		
		Write_LogEntry -Message "Explorer-Fenster geschlossen (Shell läuft weiter)." -Level "DEBUG"
	} catch {
		Write_LogEntry -Message "Fehler beim Schließen von Explorer-Fenstern: $($_)" -Level "WARNING"
	}
}

# desktop / start menu cleanup as before
$desktopShortcut = "$env:PUBLIC\Desktop\Google Chrome.lnk"
Write_LogEntry -Message "Überprüfe Desktop-Shortcut (public): $($desktopShortcut)" -Level "DEBUG"
if (Test-Path $desktopShortcut) {
    Write-Host "    Desktopeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $desktopShortcut -Force -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Entfernt: $($desktopShortcut)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Desktop-Shortcut (public) nicht vorhanden: $($desktopShortcut)" -Level "DEBUG"
}

$desktopShortcutUser = "$env:USERPROFILE\Desktop\Google Chrome.lnk"
Write_LogEntry -Message "Überprüfe Desktop-Shortcut (user): $($desktopShortcutUser)" -Level "DEBUG"
if (Test-Path $desktopShortcutUser) {
    Write-Host "    Desktopeintrag (user) vorhanden (nicht entfernt per Script)." -foregroundcolor "Cyan"
    Write_LogEntry -Message "User Desktop-Shortcut vorhanden (nicht entfernt per Script): $($desktopShortcutUser)" -Level "INFO"
} else {
    Write_LogEntry -Message "User Desktop-Shortcut nicht vorhanden: $($desktopShortcutUser)" -Level "DEBUG"
}

$startMenuShortcut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
Write_LogEntry -Message "Überprüfe Startmenu Shortcut (all users): $($startMenuShortcut)" -Level "DEBUG"
if (Test-Path $startMenuShortcut) {
    Write-Host "    Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $startMenuShortcut -Force -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Entfernt: $($startMenuShortcut)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Startmenu Shortcut (all users) nicht vorhanden: $($startMenuShortcut)" -Level "DEBUG"
}

$startMenuShortcutUser = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
Write_LogEntry -Message "Überprüfe Startmenu Shortcut (current user): $($startMenuShortcutUser)" -Level "DEBUG"
if (Test-Path $startMenuShortcutUser) {
    Write-Host "    Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $startMenuShortcutUser -Force -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Entfernt: $($startMenuShortcutUser)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Startmenu Shortcut (current user) nicht vorhanden: $($startMenuShortcutUser)" -Level "DEBUG"
}

$chromeAppsFolder = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome-Apps"
Write_LogEntry -Message "Überprüfe Chrome-Apps Ordner: $($chromeAppsFolder)" -Level "DEBUG"
if (Test-Path $chromeAppsFolder) {
    Write-Host "    Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $chromeAppsFolder -Recurse -Force -ErrorAction SilentlyContinue
    Write_LogEntry -Message "Entfernt: $($chromeAppsFolder)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Chrome-Apps Ordner nicht vorhanden: $($chromeAppsFolder)" -Level "DEBUG"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
