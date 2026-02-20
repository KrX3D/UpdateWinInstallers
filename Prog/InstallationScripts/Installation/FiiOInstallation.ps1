param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "FiiO"
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

Write_LogEntry -Message "InstallationFlag = true, starte FiiO-Installation" -Level "INFO"
Write-Host "FiiO wird installiert" -foregroundcolor "magenta"

$msiFiles = Get-ChildItem -Path "$Serverip\Daten\Treiber\FiiO_Verstarker\FiiO_v*.msi" -ErrorAction SilentlyContinue
Write_LogEntry -Message "Gefundene MSI-Dateien: $($($msiFiles).Count) im Pfad '$($Serverip)\Daten\Treiber\FiiO_Verstarker\'" -Level "DEBUG"
foreach ($msiFile in $msiFiles) {
	Write_LogEntry -Message "Starte MSI-Installation: $($msiFile.FullName)" -Level "INFO"
	[void](Invoke-InstallerFile -FilePath $msiFile.FullName -Arguments "/qn", "/passive", "/norestart" -Wait)
	Write_LogEntry -Message "MSI-Installation beendet: $($msiFile.FullName)" -Level "SUCCESS"
}

$wildcard = "FiiO_USB*"
$dpinstFiles = Get-ChildItem -Path "$Serverip\Daten\Treiber\FiiO_Verstarker\$wildcard\dpinst64.exe" -Recurse -ErrorAction SilentlyContinue
Write_LogEntry -Message "Gefundene dpinst64.exe Dateien: $($($dpinstFiles).Count)" -Level "DEBUG"
Write-Host "FiiO USB Treiber wird installiert" -foregroundcolor "magenta"
foreach ($dpinstFile in $dpinstFiles) {
	Write_LogEntry -Message "Starte dpinst: $($dpinstFile.FullName)" -Level "INFO"
	[void](Invoke-InstallerFile -FilePath $dpinstFile.FullName -Arguments "/S", "/F", "/SE", "/SA", "/SW" -Wait)
	Write_LogEntry -Message "dpinst beendet: $($dpinstFile.FullName)" -Level "SUCCESS"
}

$controlPanelLinkPath = "C:\ProgramData\Microsoft\Windows\Start Menu\FiiO\FiiO Portable High-Res Music Player series\FiiO Control Panel.lnk"
$programsDestination = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"

if (Test-Path -Path $controlPanelLinkPath) {
	Write_LogEntry -Message "ControlPanel-Link gefunden: $($controlPanelLinkPath). Verschiebe nach: $($programsDestination)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird verschoben." -foregroundcolor "Cyan"
	Move-Item -Path $controlPanelLinkPath -Destination $programsDestination -Force
	Write_LogEntry -Message "ControlPanel-Link verschoben: $($controlPanelLinkPath) -> $($programsDestination)" -Level "SUCCESS"
}

$fioDirectoryPath = "C:\ProgramData\Microsoft\Windows\Start Menu\FiiO"
$programsFioDirectoryPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\FiiO"

if (Test-Path -Path $fioDirectoryPath) {
	Write_LogEntry -Message "FiiO Startmenü-Ordner gefunden: $($fioDirectoryPath). Entferne." -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
	Remove-Item -Path $fioDirectoryPath -Recurse -Force
	Write_LogEntry -Message "FiiO Startmenü-Ordner entfernt: $($fioDirectoryPath)" -Level "SUCCESS"
}

if (Test-Path -Path $programsFioDirectoryPath) {
	Write_LogEntry -Message "Programs FiiO Ordner gefunden: $($programsFioDirectoryPath). Entferne." -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
	Remove-Item -Path $programsFioDirectoryPath -Recurse -Force
	Write_LogEntry -Message "Programs FiiO Ordner entfernt: $($programsFioDirectoryPath)" -Level "SUCCESS"
}

# ===== XML Configuration Update =====
Write_LogEntry -Message "Überprüfe und aktualisiere FiiO XML-Konfiguration..." -Level "INFO"
$xmlPath = "C:\Program Files\FiiO\FiiO_Driver\x64\FiiOCplApp.xml"

if (Test-Path -Path $xmlPath) {
	Write_LogEntry -Message "XML-Konfigurationsdatei gefunden: $xmlPath" -Level "INFO"
	
	try {
		# Load XML as text to preserve exact formatting
		$xmlText = Get-Content -Path $xmlPath -Raw -Encoding UTF8
		Write_LogEntry -Message "XML-Datei erfolgreich geladen" -Level "SUCCESS"
		
		# Also load as XML object to extract version
		[xml]$xmlContent = $xmlText
		$xmlVersion = $xmlContent.ControlPanel.PageAbout.TableDriverInfo.CustomVersion
		Write_LogEntry -Message "Version aus XML extrahiert: $xmlVersion" -Level "INFO"
		Write-Host "	Version aus XML-Datei: $xmlVersion" -foregroundcolor "Cyan"
		
		# Find and replace all <Visibility>Hidden</Visibility> with text replacement
		$modified = $false
		$changedElements = @()
		
		# Use regex to find all <Visibility>Hidden</Visibility> patterns
		$pattern = '(<Visibility>)Hidden(</Visibility>)'
		$matches = [regex]::Matches($xmlText, $pattern)
		
		if ($matches.Count -gt 0) {
			Write_LogEntry -Message "Gefundene 'Hidden' Visibility-Elemente: $($matches.Count)" -Level "INFO"
			
			# Replace all occurrences
			$xmlText = [regex]::Replace($xmlText, $pattern, '${1}Visible${2}')
			$modified = $true
			
			# For logging, determine which elements were changed
			[xml]$xmlCheck = $xmlText
			$visibilityNodes = $xmlCheck.SelectNodes("//Visibility[text()='Visible']")
			foreach ($node in $visibilityNodes) {
				$changedElements += $node.ParentNode.Name
			}
			
			Write_LogEntry -Message "Alle 'Hidden' Werte zu 'Visible' geändert" -Level "INFO"
		}
		
		# Save modified XML if changes were made
		if ($modified) {
			# Save with UTF8 encoding, preserving exact formatting
			[System.IO.File]::WriteAllText($xmlPath, $xmlText, [System.Text.Encoding]::UTF8)
			
			Write_LogEntry -Message "XML-Datei mit Änderungen gespeichert ($($matches.Count) Elemente geändert)" -Level "SUCCESS"
			Write-Host "	XML-Konfiguration aktualisiert: $($matches.Count) Element(e) von Hidden -> Visible" -foregroundcolor "Green"
			
			# Show unique changed elements
			$uniqueElements = $changedElements | Select-Object -Unique
			foreach ($element in $uniqueElements) {
				Write-Host "	  - $element" -foregroundcolor "DarkGray"
			}
		} else {
			Write_LogEntry -Message "Keine 'Hidden' Visibility-Elemente gefunden - keine Änderungen erforderlich" -Level "INFO"
		}
		
		# ===== Registry Version Correction =====
		Write_LogEntry -Message "Prüfe Registry-Versionen..." -Level "INFO"
		Write-Host "	Prüfe Registry-Einträge..." -foregroundcolor "Cyan"
		
		$RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		
		$registryEntries = foreach ($RegPath in $RegistryPaths) {
			if (Test-Path $RegPath) {
				Write_LogEntry -Message "Prüfe Registry-Pfad: $RegPath" -Level "DEBUG"
				Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "*FiiO*" -or $_.DisplayName -like "*USB DAC*" }
			}
		}
		
		if ($registryEntries) {
			foreach ($entry in $registryEntries) {
				$registryVersion = $entry.DisplayVersion
				$registryPath = $entry.PSPath
				
				Write_LogEntry -Message "Registry-Eintrag gefunden: $($entry.DisplayName), Version: $registryVersion" -Level "INFO"
				Write-Host "	Registry gefunden: $($entry.DisplayName)" -foregroundcolor "Cyan"
				Write-Host "	  Registry-Version: $registryVersion" -foregroundcolor "Cyan"
				
				# Compare versions
				if ($registryVersion -ne $xmlVersion) {
					Write_LogEntry -Message "Registry-Version ($registryVersion) stimmt nicht mit XML-Version ($xmlVersion) überein - aktualisiere Registry" -Level "WARNING"
					Write-Host "	  Versions-Diskrepanz erkannt! Aktualisiere Registry..." -foregroundcolor "Yellow"
					
					try {
						# Update DisplayVersion in registry
						Set-ItemProperty -Path $registryPath -Name "DisplayVersion" -Value $xmlVersion -ErrorAction Stop
						Write_LogEntry -Message "Registry-Version erfolgreich aktualisiert auf: $xmlVersion" -Level "SUCCESS"
						Write-Host "	  Registry-Version aktualisiert: $registryVersion -> $xmlVersion" -foregroundcolor "Green"
					} catch {
						Write_LogEntry -Message "Fehler beim Aktualisieren der Registry-Version: $_" -Level "ERROR"
						Write-Host "	  Fehler beim Aktualisieren der Registry-Version!" -foregroundcolor "Red"
					}
				} else {
					Write_LogEntry -Message "Registry-Version stimmt mit XML-Version überein: $xmlVersion" -Level "INFO"
					Write-Host "	  Registry-Version ist korrekt." -foregroundcolor "Green"
				}
			}
		} else {
			Write_LogEntry -Message "Kein FiiO-Registry-Eintrag gefunden" -Level "WARNING"
			Write-Host "	Kein FiiO-Registry-Eintrag gefunden." -foregroundcolor "Yellow"
		}
		
	} catch {
		Write_LogEntry -Message "Fehler beim Verarbeiten der XML-Datei: $_" -Level "ERROR"
		Write-Host "	Fehler beim Verarbeiten der XML-Datei!" -foregroundcolor "Red"
	}
} else {
	Write_LogEntry -Message "XML-Konfigurationsdatei nicht gefunden: $xmlPath" -Level "WARNING"
	Write-Host "	XML-Konfigurationsdatei nicht gefunden." -foregroundcolor "Yellow"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
