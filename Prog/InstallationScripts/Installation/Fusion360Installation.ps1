param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB config dateien zu kopiere. Damit Config Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Autodesk Fusion"
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
    exit
}

$InstallationFolderFusion = "$InstallationFolder\3D"
Write_LogEntry -Message "InstallationFolderFusion gesetzt: $($InstallationFolderFusion)" -Level "DEBUG"

$Installer = Join-Path $InstallationFolderFusion "Fusion_Client_Downloader.exe"
Write_LogEntry -Message "Installer-Pfad gesetzt: $($Installer)" -Level "DEBUG"

$ProjSaFolder = "$InstallationFolderFusion\ProjectSalvador*.msi"
Write_LogEntry -Message "ProjectSalvador Pattern: $($ProjSaFolder)" -Level "DEBUG"

function InstallFusion(){
    Write_LogEntry -Message "InstallFusion: Start" -Level "INFO"
    Write-Host ""
	Write-Host "Fusion 360 wird installiert" -foregroundcolor "magenta"

    # Check if installer exists
    if (-not (Test-Path $Installer)) {
        Write_LogEntry -Message "InstallFusion: Installer nicht gefunden: $($Installer)" -Level "ERROR"
        Write-Host "Installer nicht gefunden: $Installer" -ForegroundColor Red
        return $false
    } else {
        Write_LogEntry -Message "InstallFusion: Installer vorhanden: $($Installer)" -Level "DEBUG"
    }
    
	# Install Autodesk Fusion
    Write_LogEntry -Message "InstallFusion: Starte Installer $($Installer) mit Argument '--globalinstall' (Wait)" -Level "INFO"
    Start-Process -FilePath $Installer -ArgumentList '--globalinstall' -Wait
    Write_LogEntry -Message "InstallFusion: Installer beendet: $($Installer)" -Level "INFO"

	$localVersion = Get-LocalVersion
    Write_LogEntry -Message "InstallFusion: Lokale Version ermittelt: $($localVersion)" -Level "DEBUG"
	$regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\73e72ada57b7480280f7a6f4a289729f"
    Write_LogEntry -Message "InstallFusion: Ziel-Registry-Pfad: $($regPath)" -Level "DEBUG"

	# Desired values to set if key missing
	$desiredValues = @{
        "UninstallString" = '"wscript.exe" "' + (Join-Path $env:USERPROFILE "AppData\Local\Autodesk\webdeploy\meta\uninstall.wsf") + '" -a "73e72ada57b7480280f7a6f4a289729f" -p uninstall -s production'
		"DisplayName"    = "Autodesk Fusion"
        "DisplayIcon"    = Join-Path $env:USERPROFILE "AppData\Local\Autodesk\webdeploy\production\6a0c9611291d45bb9226980209917c3d\Fusion360.ico"
		"Version"        = "20250527145631"
		"DisplayVersion" = $localVersion
		"Publisher"      = "Autodesk, Inc."
        "ModifyPath"     = '"' + (Join-Path $env:USERPROFILE "AppData\Local\Autodesk\webdeploy\production\9818f225bdb3e57ac4cfc40df722e236b172bb0e\Fusion360.exe") + '" -serviceUtil'
		"NoModify"       = 0
		"NoRepair"       = 1
		"EstimatedSize"  = 3122061  # decimal of 0x004c4bed
	}

	if (-not (Test-Path $regPath)) {
		Write-Host "	Registry key nicht gefunden. Erstelle Schlüssel und setze Werte..." -ForegroundColor Yellow
        Write_LogEntry -Message "InstallFusion: RegistryKey nicht vorhanden: $($regPath). Erstelle Schlüssel und setze Werte." -Level "INFO"
		New-Item -Path $regPath -Force | Out-Null

		foreach ($name in $desiredValues.Keys) {
			$value = $desiredValues[$name]
			if ($name -in @("NoModify","NoRepair","EstimatedSize")) {
				Set-ItemProperty -Path $regPath -Name $name -Value $value -Type DWord -Force
                Write_LogEntry -Message "InstallFusion: Registry DWORD gesetzt: $($name) = $($value) für Pfad $($regPath)" -Level "DEBUG"
			} else {
				Set-ItemProperty -Path $regPath -Name $name -Value $value -Type String -Force
                Write_LogEntry -Message "InstallFusion: Registry String gesetzt: $($name) = $($value) für Pfad $($regPath)" -Level "DEBUG"
			}
		}

		Write-Host "		Alle Werte wurden gesetzt." -ForegroundColor Green
        Write_LogEntry -Message "InstallFusion: Alle Registry-Werte gesetzt." -Level "SUCCESS"
	} else {
		Write-Host "	Registry key existiert bereits. Keine Änderung vorgenommen." -ForegroundColor Cyan
        Write_LogEntry -Message "InstallFusion: RegistryKey existiert bereits: $($regPath). Keine Änderungen vorgenommen." -Level "INFO"
	}

	# Remove Autodesk Start Menu shortcuts
    $startMenuAutodesk = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Autodesk"
    if (Test-Path $startMenuAutodesk) {
        Write_LogEntry -Message "InstallFusion: Startmenu Pfad gefunden: $($startMenuAutodesk). Entferne." -Level "INFO"
        Write-Host "    Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
        Remove-Item -Path $startMenuAutodesk -Recurse -Force
        Write_LogEntry -Message "InstallFusion: Startmenu Pfad entfernt: $($startMenuAutodesk)" -Level "SUCCESS"
	}
	
	#Extension
	Write-Host "ProjectSalvador wird installiert" -foregroundcolor "magenta"
	Write_LogEntry -Message "InstallFusion: Suche ProjectSalvador Installer mit Pattern: $($ProjSaFolder)" -Level "DEBUG"
	# Install ProjectSalvador
    $projSaInstallers = Get-ChildItem -Path $ProjSaFolder -ErrorAction SilentlyContinue
    if ($projSaInstallers) {
		ForEach ($ProjInstaller in $projSaInstallers) {
            Write-Host "    Installiere: $($ProjInstaller.Name)" -ForegroundColor Cyan
            Write_LogEntry -Message "InstallFusion: Starte ProjectSalvador Installer: $($ProjInstaller.FullName)" -Level "INFO"
			Start-Process -FilePath $ProjInstaller.FullName -ArgumentList '/quiet /norestart' -Wait
            Write_LogEntry -Message "InstallFusion: ProjectSalvador Installation beendet: $($ProjInstaller.FullName)" -Level "SUCCESS"
		}
    } else {
        Write-Host "    Keine ProjectSalvador Installer gefunden in: $ProjSaFolder" -ForegroundColor Yellow
        Write_LogEntry -Message "InstallFusion: Keine ProjectSalvador Installer gefunden in: $($ProjSaFolder)" -Level "WARNING"
    }

	$shortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Autodesk"
	if (Test-Path -Path $shortcut -PathType Container) {
		Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
		Remove-Item -Path $shortcut -Recurse -Force
        Write_LogEntry -Message "InstallFusion: Entfernte Startmenu Verknüpfung: $($shortcut)" -Level "SUCCESS"
	}

	# Copy Neutron Platform folder
	$sourceNeutronPath = Join-Path $InstallationFolder "AutoDeskFusion\Neutron Platform"
	$destNeutronPath = Join-Path $env:APPDATA "Autodesk\Neutron Platform"

	Write_LogEntry -Message "InstallFusion: Neutron Platform Quelle: $($sourceNeutronPath)" -Level "DEBUG"
	Write_LogEntry -Message "InstallFusion: Neutron Platform Ziel: $($destNeutronPath)" -Level "DEBUG"

	if (Test-Path -Path $sourceNeutronPath) {
		Write-Host "    Kopiere Neutron Platform Ordner..." -ForegroundColor Cyan
		Write_LogEntry -Message "InstallFusion: Starte Kopieren von Neutron Platform von $($sourceNeutronPath) nach $($destNeutronPath)" -Level "INFO"
		
		try {
			# Create destination parent directory if it doesn't exist
			$destParent = Split-Path -Path $destNeutronPath -Parent
			if (-not (Test-Path -Path $destParent)) {
				New-Item -Path $destParent -ItemType Directory -Force | Out-Null
				Write_LogEntry -Message "InstallFusion: Zielverzeichnis erstellt: $($destParent)" -Level "DEBUG"
			}
			
			# Copy folder with all contents, overwriting existing files
			Copy-Item -Path $sourceNeutronPath -Destination $destNeutronPath -Recurse -Force -ErrorAction Stop
			
			Write-Host "        Neutron Platform erfolgreich kopiert." -ForegroundColor Green
			Write_LogEntry -Message "InstallFusion: Neutron Platform erfolgreich kopiert nach $($destNeutronPath)" -Level "SUCCESS"
		} catch {
			Write-Warning "    Fehler beim Kopieren von Neutron Platform: $_"
			Write_LogEntry -Message "InstallFusion: Fehler beim Kopieren von Neutron Platform: $($_)" -Level "ERROR"
		}
	} else {
		Write-Host "    Neutron Platform Quellordner nicht gefunden: $sourceNeutronPath" -ForegroundColor Yellow
		Write_LogEntry -Message "InstallFusion: Neutron Platform Quellordner nicht gefunden: $($sourceNeutronPath)" -Level "WARNING"
	}
	
    Write_LogEntry -Message "InstallFusion: Ende" -Level "INFO"
    return $true
}

# Define paths and patterns
$crashDumpPath = Join-Path $env:USERPROFILE "AppData\Local\CrashDumps"
$crashDumpPattern = "Fusion360*"
Write_LogEntry -Message "CrashDump Pfad: $($crashDumpPath); Pattern: $($crashDumpPattern)" -Level "DEBUG"

$prefetchPath = "C:\Windows\Prefetch"
$prefetchPattern = "FUSION*"
Write_LogEntry -Message "Prefetch Pfad: $($prefetchPath); Pattern: $($prefetchPattern)" -Level "DEBUG"

$foldersToRemove = @(
    "C:\ProgramData\Autodesk",
    (Join-Path $env:USERPROFILE "AppData\Local\Autodesk"),
    (Join-Path $env:USERPROFILE "AppData\Roaming\Autodesk"),
    "C:\Program Files\Autodesk",
    (Join-Path $env:USERPROFILE "AppData\Local\com.autodesk.cer"),
    (Join-Path $env:USERPROFILE "AppData\Local\Fusion360"),
    (Join-Path $env:USERPROFILE "AppData\Local\Fusion 360 CAM"),
    (Join-Path $env:USERPROFILE "AppData\Roaming\Fusion360")
)
Write_LogEntry -Message "Ordner-Liste für Bereinigung: $($foldersToRemove -join '; ')" -Level "DEBUG"

# Function to delete files matching pattern in a folder
function Remove-FilesIfExist {
    param (
        [string]$Path,
        [string]$Pattern
    )

    Write_LogEntry -Message "Remove-FilesIfExist: Prüfe Pfad: $($Path) auf Pattern: $($Pattern)" -Level "DEBUG"

    if (Test-Path $Path) {
        $files = Get-ChildItem -Path $Path -Filter $Pattern -File -ErrorAction SilentlyContinue
        if ($files) {
	        foreach ($file in $files) {
	            try {
	                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
	                Write-Host "	Gelöscht: $($file.FullName)" -ForegroundColor Green
                    Write_LogEntry -Message "Remove-FilesIfExist: Datei gelöscht: $($file.FullName)" -Level "SUCCESS"
	            } catch {
	                Write-Warning "	Fehler beim Löschen der Datei: $($file.FullName) - $_"
                    Write_LogEntry -Message "Remove-FilesIfExist: Fehler beim Löschen der Datei: $($file.FullName): $($_)" -Level "ERROR"
	            }
	        }
        } else {
            Write-Host "    Keine Dateien gefunden mit Pattern '$Pattern' in: $Path" -ForegroundColor Yellow
            Write_LogEntry -Message "Remove-FilesIfExist: Keine Dateien gefunden mit Pattern '$($Pattern)' in: $($Path)" -Level "INFO"
        }
    } else {
        Write-Host "	Pfad nicht gefunden: $Path" -ForegroundColor Yellow
        Write_LogEntry -Message "Remove-FilesIfExist: Pfad nicht gefunden: $($Path)" -Level "WARNING"
    }
}

# Function to remove folders recursively if they exist
function Remove-FoldersIfExist {
    param (
        [string[]]$Folders
    )
    foreach ($folder in $Folders) {
        Write_LogEntry -Message "Remove-FoldersIfExist: Prüfe Ordner: $($folder)" -Level "DEBUG"
        if (Test-Path $folder) {
            try {
                Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
                Write-Host "	Ordner gelöscht: $folder" -ForegroundColor Green
                Write_LogEntry -Message "Remove-FoldersIfExist: Ordner gelöscht: $($folder)" -Level "SUCCESS"
            } catch {
                Write-Warning "	Fehler beim Löschen des Ordners: $($folder) - $_"
                Write_LogEntry -Message "Remove-FoldersIfExist: Fehler beim Löschen des Ordners: $($folder): $($_)" -Level "ERROR"
            }
        } else {
            Write-Host "	Ordner nicht gefunden: $folder" -ForegroundColor Yellow
            Write_LogEntry -Message "Remove-FoldersIfExist: Ordner nicht gefunden: $($folder)" -Level "INFO"
        }
    }
}

function Get-LocalVersion {
    if (-Not (Test-Path $Installer)) {
        Write_LogEntry -Message "Get-LocalVersion: Installer nicht gefunden: $($Installer)" -Level "WARNING"
        return $null
    }
    $localVersion = (Get-Item $Installer).VersionInfo.FileVersion
    if (-not $localVersion) {
        Write-Host "Keine Versionsinformation verfügbar für: $Installer" -ForegroundColor Yellow
        Write_LogEntry -Message "Get-LocalVersion: Keine Versionsinformation verfügbar für: $($Installer)" -Level "WARNING"
        return $null
    }
    $parts = $localVersion.Split('.')
    if ($parts.Length -lt 3) {
        Write-Host "Ungültige lokale Version: $localVersion" -ForegroundColor Red
        Write_LogEntry -Message "Get-LocalVersion: Ungültige lokale Version: $($localVersion)" -Level "ERROR"
        return $null
    }
    # Reorder to match online structure: 2602.0.71
    $resultVersion = "$($parts[1]).$($parts[2]).$($parts[0])"
    Write_LogEntry -Message "Get-LocalVersion: Umgereihte Version: $($resultVersion)" -Level "DEBUG"
    return $resultVersion
}

function Test-FusionInstalled {
    $appid = '73e72ada57b7480280f7a6f4a289729f'
    $streamerdir = Join-Path "$env:ProgramFiles" "Autodesk\webdeploy\meta\streamer"
    
    Write_LogEntry -Message "Test-FusionInstalled: Prüfe Streamer-Verzeichnis: $($streamerdir)" -Level "DEBUG"
    # Check if streamer directory exists
    if (-not (Test-Path $streamerdir)) {
        Write-Host "    Fusion 360 Streamer-Verzeichnis nicht gefunden: $streamerdir" -ForegroundColor Yellow
        Write_LogEntry -Message "Test-FusionInstalled: Streamer-Verzeichnis nicht gefunden: $($streamerdir)" -Level "INFO"
        return $false
    }
    
    # Check for streamer executable
    $res = Get-ChildItem $streamerdir -ErrorAction SilentlyContinue | 
           Sort-Object -Descending | 
           Where-Object { $_.BaseName -match "^\d{14}$" } |
           ForEach-Object { Join-Path $_.FullName "streamer.exe" } | 
           Where-Object { Test-Path $_ }
    
    if ($res) {
        Write_LogEntry -Message "Test-FusionInstalled: Streamer gefunden: $($res[0])" -Level "DEBUG"
        return $true
    } else {
        Write-Host "    Keine gültige Fusion 360 Installation gefunden." -ForegroundColor Yellow
        Write_LogEntry -Message "Test-FusionInstalled: Keine gültige Fusion 360 Installation gefunden." -Level "INFO"
        return $false
    }
}

function Uninstall-Fusion {
    $appid = '73e72ada57b7480280f7a6f4a289729f'
    $installStream = 'production'
    $silentArgs = '--quiet'

    $streamerdir = Join-Path "$env:ProgramFiles" "Autodesk\webdeploy\meta\streamer"
    
    Write_LogEntry -Message "Uninstall-Fusion: Prüfe Streamer-Verzeichnis: $($streamerdir)" -Level "DEBUG"
    # Check if streamer directory exists
    if (-not (Test-Path $streamerdir)) {
        Write-Host "    Fusion 360 Streamer-Verzeichnis nicht gefunden: $streamerdir" -ForegroundColor Yellow
        Write-Host "    Überspringe Deinstallation, führe nur Bereinigung durch." -ForegroundColor Cyan
        Write_LogEntry -Message "Uninstall-Fusion: Streamer-Verzeichnis nicht gefunden, überspringe Deinstallation." -Level "WARNING"
        return $false
    }
    
    $res = Get-ChildItem $streamerdir -ErrorAction SilentlyContinue | 
           Sort-Object -Descending | 
           Where-Object { $_.BaseName -match "^\d{14}$" } |
           ForEach-Object { Join-Path $_.FullName "streamer.exe" } | 
           Where-Object { Test-Path $_ }
  
    if ($res) {
        if ($res -is [system.array]) {
            # we have an array, make it not...
            $res = $res[0]
        }

        Write-Host "    Deinstalliere Fusion 360 mit: $res" -ForegroundColor Cyan
        Write_LogEntry -Message "Uninstall-Fusion: Starte Deinstallation mit: $($res)" -Level "INFO"
        try {
	        #$uninstallargs="-g -p uninstall -s $installStream -a ${appid}"
	        #$uninstallargs="$silentArgs $uninstallargs"
	        #Start-ChocolateyProcessAsAdmin "$uninstallargs" "${res}" -validExitCodes $validExitCodes
            Start-Process "$res" -ArgumentList "-g","-p","uninstall","-s","$installStream","-a","$appid","--quiet" -Wait -ErrorAction Stop
            Start-Sleep 5
            Write-Host "    Deinstallation abgeschlossen." -ForegroundColor Green
            Write_LogEntry -Message "Uninstall-Fusion: Deinstallation erfolgreich abgeschlossen." -Level "SUCCESS"
            return $true
        } catch {
            Write-Warning "    Fehler bei der Deinstallation: $_"
            Write_LogEntry -Message "Uninstall-Fusion: Fehler bei Deinstallation: $($_)" -Level "ERROR"
            return $false
        }
    } else {
        Write-Host "    Keine gültige Fusion 360 Installation für Deinstallation gefunden." -ForegroundColor Yellow
        Write_LogEntry -Message "Uninstall-Fusion: Keine gültige Fusion 360 Installation für Deinstallation gefunden." -Level "INFO"
        return $false
    }
}

#Install if needed
if($InstallationFlag)
{
    Write_LogEntry -Message "InstallationFlag = true, starte Fusion Installation" -Level "INFO"
    Write-Host "Fusion 360 Installation gestartet..." -ForegroundColor Green
    $installResult = InstallFusion
    if ($installResult) {
        Write-Host "Fusion 360 Installation erfolgreich abgeschlossen." -ForegroundColor Green
        Write_LogEntry -Message "Fusion Installation erfolgreich." -Level "SUCCESS"
    } else {
        Write-Host "Fusion 360 Installation fehlgeschlagen." -ForegroundColor Red
        Write_LogEntry -Message "Fusion Installation fehlgeschlagen." -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "InstallationFlag = false, starte Deinstallation und Bereinigung" -Level "INFO"
    Write-Host "Fusion 360 wurde ohne InstallationFlag Parameter aufgerufen, Fusion wird deinstalliert." -foregroundcolor "Cyan"
    
    # Check if Fusion is actually installed before trying to uninstall
    if (Test-FusionInstalled) {
        Write-Host "Fusion 360 Installation erkannt. Starte Deinstallation..." -ForegroundColor Cyan
        Write_LogEntry -Message "Fusion Installation erkannt; starte Deinstallation." -Level "INFO"
        $uninstallResult = Uninstall-Fusion
        Write_LogEntry -Message "Uninstall-Fusion Ergebnis: $($uninstallResult)" -Level "DEBUG"
    } else {
        Write-Host "Keine Fusion 360 Installation gefunden. Überspringe Deinstallation." -ForegroundColor Yellow
        Write_LogEntry -Message "Keine Fusion Installation gefunden; überspringe Deinstallation." -Level "INFO"
        $uninstallResult = $false
    }

    # Always run cleanup regardless of uninstall result
    Write-Host ""
    Write-Host "Starte Bereinigung von Fusion360 Rückständen..." -ForegroundColor Cyan
    Write_LogEntry -Message "Starte Bereinigung von Fusion360 Rückständen" -Level "INFO"
    
    Write-Host ""
    Write-Host "Starte das Löschen von Fusion360 Crash Dumps..." -ForegroundColor Cyan
	Write_LogEntry -Message "Bereinigung: Lösche Crash Dumps in: $($crashDumpPath) mit Pattern: $($crashDumpPattern)" -Level "DEBUG"
	Remove-FilesIfExist -Path $crashDumpPath -Pattern $crashDumpPattern

	Write-Host ""
	Write-Host "Starte das Löschen von Fusion Prefetch Dateien..." -ForegroundColor Cyan
	Write_LogEntry -Message "Bereinigung: Lösche Prefetch Dateien in: $($prefetchPath) mit Pattern: $($prefetchPattern)" -Level "DEBUG"
	Remove-FilesIfExist -Path $prefetchPath -Pattern $prefetchPattern

	Write-Host ""
	Write-Host "Starte das Löschen von Autodesk/Fusion360 Ordnern..." -ForegroundColor Cyan
	Write_LogEntry -Message "Bereinigung: Entferne Ordner-Liste" -Level "DEBUG"
	Remove-FoldersIfExist -Folders $foldersToRemove

    Write-Host "Bereinigung abgeschlossen." -ForegroundColor Cyan
    Write_LogEntry -Message "Bereinigung abgeschlossen." -Level "INFO"

    # Now install fresh copy
    Write-Host ""
    Write-Host "Starte Neuinstallation von Fusion 360..." -ForegroundColor Green
    Write_LogEntry -Message "Starte Neuinstallation von Fusion 360" -Level "INFO"
    $installResult = InstallFusion
    
    if ($installResult) {
        Write-Host "Fusion 360 Neuinstallation erfolgreich abgeschlossen." -ForegroundColor Green
        Write_LogEntry -Message "Fusion Neuinstallation erfolgreich." -Level "SUCCESS"
    } else {
        Write-Host "Fusion 360 Neuinstallation fehlgeschlagen." -ForegroundColor Red
        Write_LogEntry -Message "Fusion Neuinstallation fehlgeschlagen." -Level "ERROR"
    }
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
