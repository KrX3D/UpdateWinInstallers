param(
    [string]$DriversToUpdate # Serialisierte String-Daten
)

$ProgramName = "AMD PC Treiber"
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

Write_LogEntry -Message "Script gestartet mit DriversToUpdate: $($DriversToUpdate)" -Level "INFO"
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
    Write_LogEntry -Message "Konfigurationsdatei gefunden und importiert: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

# Objekte wiederherstellen
$DriversToUpdateArray = @(
    ($DriversToUpdate -split ',') |
    Where-Object { $_ -and $_.Trim().Length -gt 0 } |
    ForEach-Object {
        $parts = $_ -split '\|'
        [PSCustomObject]@{
            DriverName        = ($parts[0] -as [string]).Trim()
            InstalledVersion  = ($parts[1] -as [string]).Trim()
            DownloadedVersion = ($parts[2] -as [string]).Trim()
            DirectoryPath     = ($parts[3] -as [string]).Trim()
        }
    }
)
Write_LogEntry -Message "DriversToUpdateArray erstellt mit $($DriversToUpdateArray.Count) Einträgen" -Level "DEBUG"

# Treiberaktionen definieren (Annahmen gemäß Ihrer Konfiguration)
$driverActions = @(
    [PSCustomObject]@{
        DriverName = "Bluetooth Driver"
        UsePnputil = $true
    },
    [PSCustomObject]@{
        DriverName = "Audio Driver"
        Exe = "Setup.exe"
        Parameters = "-s"
        UsePnputil = $false
        ExtraActions = @(
            [PSCustomObject]@{
                Action = "MKDIR"
                Command = "C:\ProgramData\UWP"
            },
            [PSCustomObject]@{
                Action = "XCOPY"
                Command = "XCOPY DirectoryPathHERE\UWP C:\ProgramData\UWP /E /R /Y"
            },
            [PSCustomObject]@{
                Action = "SCHTASKS"
                Command = 'schtasks /create /ru system /tn "install Realtek Audio UWP Services" /tr "C:\ProgramData\UWP\AsusSetup.exe -s" /rl highest /sc onlogon'
            }
        )
    },
    [PSCustomObject]@{
        DriverName = "Chipset Driver"
        #Exe = "Chipset\AMD_Chipset_Software.exe"
        Exe = "AMD_Chipset_Software.exe"
        Parameters = "/S"
        UsePnputil = $true
    },
    [PSCustomObject]@{
        DriverName = "LAN Driver"
        Exe = "setup.exe"
        Parameters = "-s"
        UsePnputil = $true
    },
    [PSCustomObject]@{
        DriverName = "Raid Driver"
        UsePnputil = $true
    },
    [PSCustomObject]@{
        DriverName = "Graphics Driver"
        Exe = "Setup.exe"
        Parameters = "-install"
        UsePnputil = $false
    },
    [PSCustomObject]@{
        DriverName = "Wi-Fi Driver"
        UsePnputil = $true
    }
)
Write_LogEntry -Message "driverActions definiert: $($driverActions.Count) Aktionen" -Level "DEBUG"

# Gesamtanzahl der Treiber ermitteln
$totalDrivers = @($DriversToUpdateArray).Count
$counter = 0
Write_LogEntry -Message "Gesamtanzahl Treiber zu verarbeiten: $($totalDrivers)" -Level "INFO"

# Treiberschleife zur Aktualisierung
foreach ($driver in $DriversToUpdateArray) {
    # Zähler für jeden Treiber erhöhen
    $counter++
	
	Write-Host 
	Write-Host "#######################################################################"
	Write-Host 
    Write-Host "	[ $counter/$totalDrivers ] $($driver.DriverName) wird installiert" -ForegroundColor "Magenta"
    Write_LogEntry -Message "Beginne Verarbeitung Treiber [$($counter)/$($totalDrivers)]: $($driver.DriverName); DirectoryPath: $($driver.DirectoryPath)" -Level "INFO"
    
	# Aktion für den aktuellen Treiber suchen
    $driverAction = $driverActions | Where-Object { $_.DriverName -eq $driver.DriverName }
    Write_LogEntry -Message "Gefundene Aktion für $($driver.DriverName): $($driverAction -ne $null)" -Level "DEBUG"
    
	if ($driverAction.Exe) {
        $exePath = Join-Path $driver.DirectoryPath $driverAction.Exe
		Write_LogEntry -Message "Prüfe exePath: $($exePath)" -Level "DEBUG"
		
        # Überprüfen, ob die ausführbare Datei im übergeordneten Verzeichnis vorhanden ist
        if (Test-Path $exePath) {
            Write_LogEntry -Message "Ausführbare Datei gefunden: $($exePath). Starte mit Parametern: $($driverAction.Parameters)" -Level "INFO"
            Start-Process -FilePath $exePath -ArgumentList $driverAction.Parameters -NoNewWindow -Wait
            Write_LogEntry -Message "Start-Process beendet für: $($exePath)" -Level "SUCCESS"
        } else {
            # Falls nicht im übergeordneten Verzeichnis vorhanden, Unterverzeichnisse durchsuchen
            Write_LogEntry -Message "Exe nicht im Root gefunden, suche in Unterverzeichnissen: $($driver.DirectoryPath)" -Level "DEBUG"
            $subDirs = Get-ChildItem -Path $driver.DirectoryPath -Recurse -Directory
            $exeInSubDir = $subDirs | ForEach-Object {
                $exe = Join-Path $_.FullName $driverAction.Exe
                if (Test-Path $exe) {
                    return $exe
                }
            } | Select-Object -First 1
            
            if ($exeInSubDir) {
                Write_LogEntry -Message "Ausführbare Datei im Unterverzeichnis gefunden: $($exeInSubDir). Starte mit Parametern: $($driverAction.Parameters)" -Level "INFO"
                Start-Process -FilePath $exeInSubDir -ArgumentList $driverAction.Parameters -NoNewWindow -Wait
                Write_LogEntry -Message "Start-Process beendet für: $($exeInSubDir)" -Level "SUCCESS"
            } else {
                Write_LogEntry -Message "Keine ausführbare Datei gefunden für $($driver.DriverName) im Pfad $($driver.DirectoryPath)" -Level "WARNING"
            }
        }
    }
        
	# Überprüfen, ob pnputil.exe ebenfalls ausgeführt werden soll
	if ($driverAction.UsePnputil) {
		Write_LogEntry -Message "Führe pnputil für $($driver.DriverName) aus in $($driver.DirectoryPath)" -Level "INFO"
		#Write-Host "pnputil.exe wird ausgeführt für $($driver.DriverName) unter $($driver.DirectoryPath)"
		Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "pnputil.exe /add-driver $($driver.DirectoryPath)\*.inf /install /subdirs > nul 2>&1" -NoNewWindow -Wait
		Write_LogEntry -Message "pnputil-Aufruf beendet für $($driver.DriverName)" -Level "SUCCESS"
	}

	# Zusätzliche Aktionen ausführen (Allgemein für alle Treiber)
	if ($driverAction.ExtraActions) {
		foreach ($extraAction in $driverAction.ExtraActions) {
			Write_LogEntry -Message "Führe ExtraAction $($extraAction.Action) aus für $($driver.DriverName)" -Level "DEBUG"
			switch ($extraAction.Action) {
				"MKDIR" {
					#Write-Host "Erstelle Verzeichnis: $($extraAction.Command)"
					Write_LogEntry -Message "Erstelle Verzeichnis: $($extraAction.Command)" -Level "INFO"
					New-Item -ItemType Directory -Force -Path $extraAction.Command | Out-Null
					Write_LogEntry -Message "Verzeichnis erstellt: $($extraAction.Command)" -Level "SUCCESS"
				}
				"XCOPY" {
					$commandToRun = $extraAction.Command -replace 'DirectoryPathHERE', $driver.DirectoryPath
					Write_LogEntry -Message "Starte XCOPY: $($commandToRun)" -Level "INFO"
					# XCOPY in einem Job ausführen, um Abschluss sicherzustellen
					$job = Start-Job -ScriptBlock { Invoke-Expression $using:commandToRun }
					Wait-Job $job | Out-Null
					Remove-Job $job | Out-Null
					Write_LogEntry -Message "XCOPY abgeschlossen: $($commandToRun)" -Level "SUCCESS"
					#Invoke-Expression $commandToRun
				}
				"SCHTASKS" {
					Write_LogEntry -Message "Erstelle geplante Aufgabe: $($extraAction.Command)" -Level "INFO"
					# SCHTASKS in einem Job ausführen, um Abschluss sicherzustellen
					$job = Start-Job -ScriptBlock { Invoke-Expression $using:extraAction.Command }
					Wait-Job $job | Out-Null
					Remove-Job $job | Out-Null
					Write_LogEntry -Message "Scheduled Task erstellt: $($extraAction.Command)" -Level "SUCCESS"
					#Invoke-Expression $extraAction.Command
				}
				default {
					Write_LogEntry -Message "Unbekannte ExtraAction: $($extraAction.Action) für $($driver.DriverName)" -Level "WARNING"
					Write-Host "Unbekannte Aktion: $($extraAction.Action)"
				}
			}
		}
	}
	Write_LogEntry -Message "Verarbeitung abgeschlossen für Treiber: $($driver.DriverName) [$($counter)/$($totalDrivers)]" -Level "INFO"
}

Write-Host 
Write-Host "#######################################################################"
Write_LogEntry -Message "Alle Treiber verarbeitet. Gesamt: $($totalDrivers)" -Level "SUCCESS"

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
