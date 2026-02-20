param(
    [switch]$InstallationFlag, #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
    [switch]$InstallVirtBox,
    [switch]$InstallVirtBoxExPack
)

$ProgramName = "Oracle VirtualBox"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag), InstallVirtBox: $($InstallVirtBox), InstallVirtBoxExPack: $($InstallVirtBoxExPack)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

#Bei Update wird ohne deinstallation die neue Version richtig installiert

if ($InstallationFlag -eq $true) {
	Write-Host "Microsoft Visual C++ wird installiert" -foregroundcolor "magenta"
	Write_LogEntry -Message "Installiere Microsoft Visual C++ (InstallationFlag true)." -Level "INFO"

	# Install Microsoft Visual C++
	$vcInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\VirtualBox\VC*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
	Write_LogEntry -Message "Visual C++ Installer gesucht: $($Serverip)\Daten\Prog\VirtualBox\VC*.exe; Gefunden: $($vcInstaller.FullName)" -Level "DEBUG"
	if ($vcInstaller) {
		Write_LogEntry -Message "Starte Visual C++ Installer: $($vcInstaller.FullName)" -Level "INFO"
		[void](Invoke-InstallerFile -FilePath $vcInstaller.FullName -Arguments '/install', '/passive', '/qn', '/norestart' -Wait)
        Write_LogEntry -Message "Visual C++ Installer beendet: $($vcInstaller.FullName)" -Level "SUCCESS"
	}
}

# Install VirtualBox
$virtualBoxInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\VirtualBox*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
Write_LogEntry -Message "VirtualBox Installer gesucht: $($Serverip)\Daten\Prog\VirtualBox*.exe; Gefunden: $($virtualBoxInstaller.FullName)" -Level "DEBUG"

if (($InstallationFlag -eq $true) -or (($virtualBoxInstaller) -and ($InstallVirtBox -eq $true))) {
	Write-Host "VirtualBox wird installiert" -foregroundcolor "magenta"
    Write_LogEntry -Message "VirtualBox Installationsbedingung erfüllt. Starte Installer: $($virtualBoxInstaller.FullName)" -Level "INFO"
	[void](Invoke-InstallerFile -FilePath $virtualBoxInstaller.FullName -Arguments '--silent', '--ignore-reboot', '--msiparams "VBOX_INSTALLQUICKLAUNCHSHORTCUT=0"' -Wait)
    Write_LogEntry -Message "VirtualBox Installer ausgeführt: $($virtualBoxInstaller.FullName)" -Level "SUCCESS"
	
	$InstallationFileVirtualBox = "C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"
	$virtualBoxInstaller = Get-ChildItem -Path $InstallationFileVirtualBox | Select-Object -Last 1
    Write_LogEntry -Message "Gefundene VirtualBox Executable: $($InstallationFileVirtualBox); Gefunden: $($virtualBoxInstaller.FullName)" -Level "DEBUG"
	$localVersion = (Get-Item $virtualBoxInstaller.FullName).VersionInfo.ProductVersion
	$versionParts = $localVersion -split '\.'
	$localVersion = $versionParts[-1]
	Write_LogEntry -Message "Extracted VirtualBox localVersion part: $($localVersion)" -Level "DEBUG"
	
	$EstlPath  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }
	Write_LogEntry -Message "Registry-Eintrag für ProgramName gesucht: $($ProgramName); Gefunden: $($EstlPath.PSPath)" -Level "DEBUG"

	$RegistryPath = $EstlPath.PSPath  # Get the Registry path from $EstlPath
	Write_LogEntry -Message "RegistryPath zur Aktualisierung: $($RegistryPath)" -Level "DEBUG"
	# Create or update the DisplayVersion value in the Registry
	Set-ItemProperty -Path $RegistryPath -Name 'VersionRevision' -Value $localVersion -Type DWord
	Write_LogEntry -Message "Registry 'VersionRevision' gesetzt: $($localVersion) in $($RegistryPath)" -Level "INFO"
}

# Install VirtualBox Extension Pack
#$extensionPack = Get-ChildItem -Path "$Serverip\Daten\Prog\Oracle_VM_VirtualBox_Extension_Pack*.vbox-extpack" -ErrorAction SilentlyContinue | Select-Object -First 1
$extensionPack = Get-ChildItem -Path "$Serverip\Daten\Prog\Oracle_VirtualBox_Extension_Pack*.vbox-extpack" -ErrorAction SilentlyContinue | Select-Object -First 1
Write_LogEntry -Message "Gefundene Extension Pack Datei: $($extensionPack.FullName)" -Level "DEBUG"

if (($InstallationFlag -eq $true) -or (($extensionPack) -and ($InstallVirtBoxExPack -eq $true))) {
	Write-Host "VirtualBox Extensions wird installiert" -foregroundcolor "magenta"
    Write_LogEntry -Message "Installiere VirtualBox Extension Pack: $($extensionPack.FullName)" -Level "INFO"
	Set-Alias vboxmanage "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
	"y" | vboxmanage extpack install --replace $extensionPack.FullName
	if ($LastExitCode -ne 0) {
		Write_Warning "Extension pack installation failed with exit code $LastExitCode"
        Write_LogEntry -Message "Extension pack installation failed with exit code $($LastExitCode)" -Level "ERROR"
	} else {
        Write_LogEntry -Message "Extension pack installation completed successfully: $($extensionPack.FullName)" -Level "SUCCESS"
    }
}

# Check if Oracle VM VirtualBox Start Menu shortcuts exist
if (Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Oracle VirtualBox") {
    # Remove Oracle VM VirtualBox Start Menu shortcuts
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Oracle VirtualBox" -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Oracle VirtualBox" -Level "INFO"
}

# Check if Oracle VM VirtualBox Quick Launch shortcut exists
$quickLaunchShortcut = Join-Path $env:AppData "Microsoft\Internet Explorer\Quick Launch\Oracle VirtualBox.lnk"
Write_LogEntry -Message "Überprüfe QuickLaunch-Shortcut: $($quickLaunchShortcut)" -Level "DEBUG"
if (Test-Path $quickLaunchShortcut) {
    # Delete Oracle VM VirtualBox Quick Launch shortcut
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $quickLaunchShortcut -Force
    Write_LogEntry -Message "QuickLaunch-Shortcut entfernt: $($quickLaunchShortcut)" -Level "INFO"
}

if ($InstallationFlag -eq $true) {
	# Ändere Interface-Einstellungen
	$xmlFilePath = "$env:USERPROFILE\.VirtualBox\VirtualBox.xml"
	Write_LogEntry -Message "VirtualBox XML Pfad: $($xmlFilePath)" -Level "DEBUG"
	
	# Überprüfen, ob die Datei existiert
	if (-not (Test-Path $xmlFilePath)) {
		# Starte VirtualBox, um die XML-Datei zu generieren
		Write-Host 
		Write-Host "Die Datei VirtualBox.xml wurde nicht gefunden. VirtualBox wird gestartet, um sie zu erstellen..." -ForegroundColor Yellow
		Write_LogEntry -Message "VirtualBox.xml nicht gefunden, starte VirtualBox zum Erzeugen der Datei." -Level "WARNING"
		
		# Pfad zur VirtualBox.exe angeben
		$virtualBoxPath = "C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"
		Write_LogEntry -Message "Erwarteter VirtualBox Pfad: $($virtualBoxPath)" -Level "DEBUG"

		# Überprüfen, ob die Datei existiert
		if (-not (Test-Path $virtualBoxPath)) {
			Write-Host "Die VirtualBox.exe wurde am erwarteten Speicherort nicht gefunden: $virtualBoxPath" -ForegroundColor Red
            Write_LogEntry -Message "VirtualBox.exe nicht gefunden: $($virtualBoxPath)" -Level "ERROR"
			return
		}

		# Starte VirtualBox
		$virtualBoxProcess = Start-Process -FilePath $virtualBoxPath -PassThru
        Write_LogEntry -Message "VirtualBox Prozess gestartet zum Erzeugen der XML-Datei. PID: $($virtualBoxProcess.Id)" -Level "DEBUG"

		# Warte 3 Sekunden, um sicherzustellen, dass die Datei erstellt wird
		Start-Sleep -Seconds 3

		# Versuche, VirtualBox normal zu schließen
		Write-Host "Schließe VirtualBox..."
		$virtualBoxProcess.CloseMainWindow() | Out-Null

		# Warte, bis der Prozess beendet ist
		$virtualBoxProcess.WaitForExit()

		# Überprüfen, ob die Datei jetzt existiert
		if (Test-Path $xmlFilePath) {
			Write-Host "Die Datei VirtualBox.xml wurde erfolgreich erstellt." -ForegroundColor Green
            Write_LogEntry -Message "VirtualBox.xml erfolgreich erstellt: $($xmlFilePath)" -Level "SUCCESS"
		} else {
			Write-Host "Die Datei VirtualBox.xml wurde nach dem Start von VirtualBox nicht erstellt." -ForegroundColor Red
            Write_LogEntry -Message "VirtualBox.xml konnte nach Start von VirtualBox nicht erstellt werden: $($xmlFilePath)" -Level "ERROR"
			return
		}
	}

	# Überprüfen, ob die Datei existiert
	if (Test-Path $xmlFilePath) {
		# XML-Inhalt laden
		[xml]$xmlContent = Get-Content -Path $xmlFilePath
        Write_LogEntry -Message "VirtualBox.xml geladen: $($xmlFilePath)" -Level "DEBUG"
		
		# Überprüfen, ob der ExtraDataItem für GUI/ColorTheme existiert
		$colorThemeItem = $xmlContent.VirtualBox.Global.ExtraData.ExtraDataItem | Where-Object { $_.name -eq "GUI/ColorTheme" }

		if ($colorThemeItem) {
			# Falls der ExtraDataItem existiert, ändere den Wert auf "Dark"
			$colorThemeItem.value = "Dark"
            Write_LogEntry -Message "GUI/ColorTheme existierte und wurde auf 'Dark' gesetzt." -Level "INFO"
		} else {
			# Falls der ExtraDataItem nicht existiert, erstelle einen neuen mit dem Wert "Dark"
			$newColorThemeItem = $xmlContent.CreateElement("ExtraDataItem")
			$newColorThemeItem.SetAttribute("name", "GUI/ColorTheme")
			$newColorThemeItem.SetAttribute("value", "Dark")
			
			# Füge den neuen ExtraDataItem am Anfang des ExtraData-Abschnitts ein
			$xmlContent.VirtualBox.Global.ExtraData.InsertBefore($newColorThemeItem, $xmlContent.VirtualBox.Global.ExtraData.FirstChild)
            Write_LogEntry -Message "GUI/ColorTheme neu erstellt und auf 'Dark' gesetzt." -Level "INFO"
		}

		# Verarbeiten von GUI/SuppressMessages
		$suppressMessagesItem = $xmlContent.VirtualBox.Global.ExtraData.ExtraDataItem | Where-Object { $_.name -eq "GUI/SuppressMessages" }
		$newSuppressValue = "confirmResetMachine,confirmGoingFullscreen,remindAboutMouseIntegration,remindAboutAutoCapture"

		if ($suppressMessagesItem) {
			# Falls der Schlüssel existiert, aktualisiere seinen Wert, um die erforderlichen Nachrichten hinzuzufügen
			$currentValue = $suppressMessagesItem.value
			$updatedValue = ($currentValue.Split(',') + $newSuppressValue.Split(',') | Sort-Object -Unique) -join ','
			$suppressMessagesItem.value = $updatedValue
            Write_LogEntry -Message "GUI/SuppressMessages aktualisiert: $($updatedValue)" -Level "DEBUG"
		} else {
			# Falls der Schlüssel fehlt, erstelle ihn
			$newSuppressItem = $xmlContent.CreateElement("ExtraDataItem")
			$newSuppressItem.SetAttribute("name", "GUI/SuppressMessages")
			$newSuppressItem.SetAttribute("value", $newSuppressValue)
			$xmlContent.VirtualBox.Global.ExtraData.AppendChild($newSuppressItem)
            Write_LogEntry -Message "GUI/SuppressMessages neu erstellt: $($newSuppressValue)" -Level "INFO"
		}

		# Entfernen der GUI/NotificationCenter-Schlüssel
		$notificationKeys = $xmlContent.VirtualBox.Global.ExtraData.ExtraDataItem | Where-Object { $_.name -like "GUI/NotificationCenter*" }
		foreach ($key in $notificationKeys) {
			$key.ParentNode.RemoveChild($key) | Out-Null
            Write_LogEntry -Message "GUI/NotificationCenter Schlüssel entfernt: $($key.name)" -Level "DEBUG"
		}
		
		# Speichere den bearbeiteten XML-Inhalt zurück in die Datei
		$xmlContent.Save($xmlFilePath)
        Write_LogEntry -Message "VirtualBox.xml gespeichert: $($xmlFilePath)" -Level "SUCCESS"

		# Entferne manuell alle xmlns="" aus der Datei als letzter Schritt
		$xmlContentAsString = Get-Content -Path $xmlFilePath -Raw
		$xmlContentAsString = $xmlContentAsString -replace ' xmlns=""', ''

		# Entferne unnötige Leerzeichen vor "/>" (Aufräumen)
		$xmlContentAsString = $xmlContentAsString -replace ' />', '/>'

		# Schreibe die finale Version zurück in die Datei
		Set-Content -Path $xmlFilePath -Value $xmlContentAsString
        Write_LogEntry -Message "VirtualBox.xml final bereinigt und geschrieben: $($xmlFilePath)" -Level "DEBUG"
		
		#Write-Host "Die VirtualBox-Einstellungen wurden erfolgreich aktualisiert." -ForegroundColor Green
	} else {
		Write-Host "" 
		Write-Host "Die Datei VirtualBox.xml existiert nicht am erwarteten Speicherort: $xmlFilePath" -ForegroundColor Yellow
        Write_LogEntry -Message "VirtualBox.xml existiert nicht nach Versuch: $($xmlFilePath)" -Level "WARNING"
	}
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
