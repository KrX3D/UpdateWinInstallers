param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "EstlCam"
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
Write_LogEntry -Message "ProgramName gesetzt: $($ProgramName); ScriptType gesetzt: $($ScriptType)" -Level "DEBUG"

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
    Write_LogEntry -Message "Script beendet wegen fehlender Konfigurationsdatei: $($configPath)" -Level "ERROR"
    Finalize_LogSession
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

$EstlcamFolder = "$Serverip\Daten\Prog\CNC\Estlcam_64*.exe"
$EstlcamDesktopCrackFile = "$Serverip\Daten\Prog\CNC\EstlCam\Estlcam.bat"

function Get-EstlcamBuildFromFilename {
    param([Parameter(Mandatory)][string]$Name)
    $m = [regex]::Match($Name, 'Estlcam_64_(\d+)\.exe', 'IgnoreCase')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return $null
}

function Get-EstlcamMainVersionFromBuild {
    param([Parameter(Mandatory)][int]$Build)
    $s = $Build.ToString()
    if ($s.Length -ge 2) { return [int]$s.Substring(0,2) }
    return [int]$s
}

function Remove-OldPublicDesktopShortcuts {
    param(
        [Parameter(Mandatory)][int]$KeepMainVersion
    )

    $publicDesktop = [Environment]::GetFolderPath('CommonDesktopDirectory')  # C:\Users\Public\Desktop
    Write_LogEntry -Message "Prüfe alte Public-Desktop Verknüpfungen in: $publicDesktop (Keep V$KeepMainVersion)" -Level "INFO"

    if (-not (Test-Path $publicDesktop)) {
        Write_LogEntry -Message "Public Desktop existiert nicht: $publicDesktop" -Level "WARNING"
        return
    }

    $lnks = Get-ChildItem -Path $publicDesktop -Filter "Estlcam V*.lnk" -ErrorAction SilentlyContinue
    foreach ($l in $lnks) {
        if ($l.Name -match '^Estlcam V(\d+)\s+(CAM|CNC)\.lnk$') {
            $v = [int]$Matches[1]
            if ($v -ne $KeepMainVersion) {
                try {
                    Remove-Item -Path $l.FullName -Force
                    Write_LogEntry -Message "Alte Verknüpfung entfernt: $($l.FullName)" -Level "SUCCESS"
                } catch {
                    Write_LogEntry -Message "Konnte Verknüpfung nicht löschen: $($l.FullName) - $($_.Exception.Message)" -Level "WARNING"
                }
            }
        }
    }
}

Write_LogEntry -Message "Estlcam-Quellmuster: $($EstlcamFolder); DesktopCrackFile: $($EstlcamDesktopCrackFile)" -Level "DEBUG"
Write-Host "Estlcam wird installiert" -foregroundcolor "magenta"

# Install Estlcam
$installerList = Get-ChildItem -Path $EstlcamFolder -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

Write_LogEntry -Message "Gefundene Estlcam-Installer im Ordner: $($installerList.Count) Elemente" -Level "INFO"

foreach ($Installer in $installerList) {
	# Create a shared flag variable
	$script:StopBrowserMonitor = $false
    Write_LogEntry -Message "Beginne Installation von: $($Installer.FullName)" -Level "INFO"

    # Start a background job to monitor and close browser windows after a delay
    $browserMonitor = Start-Job -ScriptBlock {
        while (-not $script:StopBrowserMonitor) {
			Get-Process | Where-Object { $_.Name -match 'chrome|msedge' } | ForEach-Object {
            #Get-Process -Name chrome, msedge -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Host "Schließe Browser: $($_.Name)" -ForegroundColor "cyan"
                Stop-Process -Id $_.Id -Force
            }
            Start-Sleep -Seconds 1
        }
    }
    Write_LogEntry -Message "Browser-Monitor Job gestartet: Id=$($browserMonitor.Id)" -Level "DEBUG"

    Write_LogEntry -Message "Starte Estlcam Installer: $($Installer.FullName) mit Argument '/S' (synchronously)" -Level "INFO"
    Start-Process -FilePath $Installer.FullName -ArgumentList '/S' -Wait
    Write_LogEntry -Message "Installer-Prozess beendet für: $($Installer.FullName)" -Level "SUCCESS"

	# Signal the job to stop gracefully
	$script:StopBrowserMonitor = $true
    Write_LogEntry -Message "Signal gesendet: StopBrowserMonitor = $($script:StopBrowserMonitor)" -Level "DEBUG"

	# Wait up to 10 seconds for the job to stop gracefully
	Write_LogEntry -Message "Warte auf Beendigung des Browser-Monitor Jobs (Timeout 10s)" -Level "DEBUG"
	Wait-Job -Job $browserMonitor -Timeout 10

	# If the job is still running, forcefully stop it
	if ($browserMonitor.State -eq 'Running') {
		Write-Host "Hintergrundprozess wird erzwungen gestoppt." -ForegroundColor "Red"
        Write_LogEntry -Message "Browser-Monitor Job läuft noch nach Timeout; stoppe Job Id=$($browserMonitor.Id)" -Level "WARNING"

        # Stop job now (even if it would auto-end)
	    try {
	        Stop-Job -Job $browserMonitor -ErrorAction SilentlyContinue
	        Write_LogEntry -Message "Stop-Job aufgerufen für Job Id=$($browserMonitor.Id)" -Level "DEBUG"
	    } catch {
	        Write_LogEntry -Message "Browser-Monitor Job cleanup failed: $($_.Exception.Message)" -Level "WARNING"
	    }
	}

	# Ensure the job is removed
	Remove-Job -Job $browserMonitor
	Write-Host "Browser Monitor Job entfernt." -ForegroundColor "Green"
    Write_LogEntry -Message "Browser-Monitor Job entfernt: Id=$($browserMonitor.Id)" -Level "SUCCESS"
	
    $InstallationFileName = $Installer.Name
    Write_LogEntry -Message "Installationsdatei Name: $($InstallationFileName)" -Level "DEBUG"

    $localVersion = Get-EstlcamBuildFromFilename -Name $InstallationFileName
	Write_LogEntry -Message "Lokale Version extrahiert: $($localVersion)" -Level "DEBUG"
	

    $localMainVersion = Get-EstlcamMainVersionFromBuild -Build $localVersion
	Write_LogEntry -Message "Lokale Main-Version extrahiert: $($localMainVersion)" -Level "DEBUG"
		
    $EstlcamVFolder = "$Serverip\Daten\Prog\CNC\EstlCam\V$localMainVersion"
    $destination    = "C:\ProgramData\Estlcam\V$localMainVersion"
    Write_LogEntry -Message "Estlcam V-Folder: $EstlcamVFolder; Destination: $destination" -Level "DEBUG"

	$EstlPath  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }
	Write_LogEntry -Message "Registry-Abfrage durchgeführt für ProgramName-Muster: $($ProgramName + '*')" -Level "DEBUG"
	Write_LogEntry -Message "Gefundene Registry-Einträge: $($($EstlPath).Count)" -Level "DEBUG"

	$RegistryPath = $EstlPath.PSPath  # Get the Registry path from $EstlPath
	Write_LogEntry -Message "RegistryPath ermittelt: $($RegistryPath)" -Level "DEBUG"
	
	# Create or update the DisplayVersion value in the Registry
	Write_LogEntry -Message "Setze Registry-DisplayVersion auf: $($localVersion) für Pfad: $($RegistryPath)" -Level "INFO"
	Set-ItemProperty -Path $RegistryPath -Name 'DisplayVersion' -Value $localVersion -Type String
	Write_LogEntry -Message "Registry-DisplayVersion geschrieben: $($localVersion) für Pfad: $($RegistryPath)" -Level "SUCCESS"
}

# Copy Estlcam VXX files
#If (Test-Path $EstlcamVFolder -and $InstallationFlag -eq $true) {
	#Write-Host "	Backup wird wiederhergestellt." -foregroundcolor "Cyan"
    #Copy-Item -Path $EstlcamVFolder -Destination ("C:\ProgramData\Estlcam\V" + $localMainVersion) -Recurse -Force
#}

if ((Test-Path $EstlcamVFolder) -and ($InstallationFlag -eq $true)) {
	Write_LogEntry -Message "Backup-Ordner gefunden: $($EstlcamVFolder) und InstallationFlag ist true. Starte Wiederherstellung." -Level "INFO"
	Write-Host "	Backup wird wiederhergestellt." -foregroundcolor "Cyan"
	if (!(Test-Path $destination)) {
		Write_LogEntry -Message "Ziel existiert nicht: $($destination) - erstelle Verzeichnis" -Level "DEBUG"
		New-Item -ItemType Directory -Path $destination -Force | Out-Null
	}
    try {
        Get-ChildItem $EstlcamVFolder -ErrorAction SilentlyContinue | Copy-Item -Destination $destination -Recurse -Force
    	Write_LogEntry -Message "Backup kopiert: Quelle=$($EstlcamVFolder) -> Ziel=$($destination)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Backup Kopieren fehlgeschlagen: $($_.Exception.Message)" -Level "WARNING"
    }
}

# Remove Estlcam Start Menu shortcuts (current main version folder)
$shortCut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Estlcam " + $localMainVersion
Write_LogEntry -Message "Prüfe Startmenü-Verknüpfung: $($shortCut)" -Level "DEBUG"
If (Test-Path $shortCut) {
    Write-Host "    Startmenüeintrag wird entfernt." -ForegroundColor "Cyan"
    try {
        Remove-Item -Path $shortCut -Recurse -Force
        Write_LogEntry -Message "Startmenüeintrag entfernt: $shortCut" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Startmenüeintrag konnte nicht entfernt werden: $($_.Exception.Message)" -Level "WARNING"
    }
}

# Remove old Public Desktop shortcuts (keep only current main version)
Remove-OldPublicDesktopShortcuts -KeepMainVersion $localMainVersion

# Create Estlcam Desktop Crack
Write_LogEntry -Message "Prüfe Desktop-Crack Datei: $($EstlcamDesktopCrackFile)" -Level "DEBUG"
If (Test-Path $EstlcamDesktopCrackFile) {
	Write-Host "	Desktop Crack wird kopiert." -foregroundcolor "Cyan"
    try {
        Copy-Item -Path $EstlcamDesktopCrackFile -Destination "$env:USERPROFILE\Desktop\" -Force
        Write_LogEntry -Message "Desktop-Crack kopiert nach: $($env:USERPROFILE)\Desktop" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Desktop-Crack Kopieren fehlgeschlagen: $($_.Exception.Message)" -Level "WARNING"
    }
} else {
    Write_LogEntry -Message "Desktop-Crack Datei nicht gefunden: $($EstlcamDesktopCrackFile)" -Level "WARNING"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
