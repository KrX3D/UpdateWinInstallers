param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "PrusaSlicer"
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

#Bei Update wird ohne deinstallation die neue Version richtig installiert // Seit 2.7 nicht mehr

$PrusaConfigFolder = "$Serverip\Daten\3D Printer\Config\PrusaSlicer\PrusaSlicer"
Write_LogEntry -Message "PrusaConfigFolder: $($PrusaConfigFolder)" -Level "DEBUG"

# Check if Prusa Slicer is installed
$uninstallCommand = "C:\Program Files\Prusa3D\PrusaSlicer\unins000.exe"
Write_LogEntry -Message "Uninstall command path: $($uninstallCommand)" -Level "DEBUG"
if (Test-Path $uninstallCommand) {
    Write_LogEntry -Message "PrusaSlicer Deinstallation: Deinstallationsprogramm gefunden: $($uninstallCommand)" -Level "INFO"
    Write-Host "PrusaSlicer ist installiert, Deinstallation beginnt." -foregroundcolor "magenta"
	Start-Process $uninstallCommand -ArgumentList "/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART" -Wait
    Write_LogEntry -Message "PrusaSlicer Deinstallationsprozess abgeschlossen: $($uninstallCommand)" -Level "SUCCESS"
    Write-Host "	PrusaSlicer wurde deinstalliert." -foregroundcolor "green"
    Start-Sleep -Seconds 3
} else {
    Write_LogEntry -Message "PrusaSlicer Deinstallation: Deinstallationsprogramm nicht gefunden: $($uninstallCommand)" -Level "DEBUG"
}

Write-Host "PrusaSlicer wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "PrusaSlicer Installationsroutine gestartet." -Level "INFO"

$PrusaPath = "$Serverip\Daten\Prog\3D\prusaslicer*.exe"
Write_LogEntry -Message "Suchmuster für Prusa Installer: $($PrusaPath)" -Level "DEBUG"
$PrusaFile = Get-ChildItem -Path $PrusaPath | Select-Object -First 1 -ExpandProperty FullName
Write_LogEntry -Message ('Gefundene Prusa Installer-Datei: ' + $($(if ($PrusaFile) { $PrusaFile } else { '<none>' }))) -Level "DEBUG"

if (Test-Path $PrusaFile) {
    Write_LogEntry -Message "Starte PrusaSlicer Installer: $($PrusaFile)" -Level "INFO"
	[void](Invoke-InstallerFile -FilePath $PrusaFile -Arguments "/exenoui /exenoupdates /silent /norestart" -Wait)
    Write_LogEntry -Message "PrusaSlicer Installer ausgeführt: $($PrusaFile)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein PrusaSlicer Installer gefunden mit Pattern $($PrusaPath)" -Level "WARNING"
}

if ((Test-Path $PrusaConfigFolder) -and ($InstallationFlag -eq $true)) {
    $destinationFolder = "$env:USERPROFILE\AppData\Roaming\PrusaSlicer"
	Write_LogEntry -Message "PrusaConfigFolder vorhanden: $($PrusaConfigFolder). Destination: $($destinationFolder)" -Level "INFO"
	#Write-Host "PrusaSlicer config wird wiederhergestellt" -foregroundcolor "Yellow"
	#Copy-Item -Path $PrusaConfigFolder -Destination $destinationFolder -Recurse -Force
	
	if (Test-Path $PrusaConfigFolder) {
		Write_LogEntry -Message "Beginne Wiederherstellung PrusaSlicer Config von $($PrusaConfigFolder) nach $($destinationFolder)" -Level "DEBUG"
		Write-Host "PrusaSlicer config wird wiederhergestellt" -foregroundcolor "Yellow"
		if (!(Test-Path $destinationFolder)) {
			New-Item -ItemType Directory -Path $destinationFolder -Force
            Write_LogEntry -Message "Zielordner erstellt: $($destinationFolder)" -Level "DEBUG"
		}
		Get-ChildItem $PrusaConfigFolder | Copy-Item -Destination $destinationFolder -Recurse -Force
        Write_LogEntry -Message "PrusaSlicer Config erfolgreich kopiert." -Level "SUCCESS"
	}
} else {
    Write_LogEntry -Message "Keine PrusaConfigFolder vorhanden oder InstallationFlag nicht gesetzt. PrusaConfigFolderExists: $([bool](Test-Path $PrusaConfigFolder)); InstallationFlag: $($InstallationFlag)" -Level "DEBUG"
}

if (Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Prusa3D") {
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Prusa3D" -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Prusa3D" -Level "INFO"
    Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Prusa G-code Viewer.lnk" -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: Prusa G-code Viewer.lnk" -Level "DEBUG"
    Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PrusaSlicer*.lnk" -Force
    Write_LogEntry -Message "Startmenüeintrag(en) PrusaSlicer*.lnk entfernt" -Level "DEBUG"
}

if (Test-Path "C:\Users\Public\Desktop\Printables - Free 3D model library.url") {
	Write-Host "	Printables - Free 3D model library URL wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "C:\Users\Public\Desktop\Printables - Free 3D model library.url" -Force
    Write_LogEntry -Message "Öffentliche Desktop-URL entfernt: Printables - Free 3D model library.url" -Level "INFO"
}

if (Test-Path "C:\Users\Public\Desktop\Prusa3D.lnk") {
	Write-Host "	Desktop Verknüpfung wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "C:\Users\Public\Desktop\Prusa3D.lnk" -Force
    Write_LogEntry -Message "Öffentliche Desktop-Verknüpfung entfernt: Prusa3D.lnk" -Level "INFO"
}

if (Test-Path "C:\Users\Public\Desktop\Prusa G-code Viewer.lnk") {
	Write-Host "	Desktop Verknüpfung wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "C:\Users\Public\Desktop\Prusa G-code Viewer.lnk" -Force
    Write_LogEntry -Message "Öffentliche Desktop-Verknüpfung entfernt: Prusa G-code Viewer.lnk" -Level "INFO"
}

# Define the path to the directory containing the shortcut
$shortcutDirectory = "C:\Users\Public\Desktop\"
Write_LogEntry -Message "Shortcut directory: $($shortcutDirectory)" -Level "DEBUG"

# Get the PrusaSlicer shortcut
$prusaSlicerShortcut = Get-ChildItem -Path $shortcutDirectory -Filter "PrusaSlicer*.lnk"
Write_LogEntry -Message ('Gefundene PrusaSlicer Shortcuts: ' + $($(if ($prusaSlicerShortcut) { $prusaSlicerShortcut.Count } else { 0 }))) -Level "DEBUG"

# Check if PrusaSlicer shortcut exists
if ($prusaSlicerShortcut) {
    # Get the shortcut name
    $shortcutName = $prusaSlicerShortcut.BaseName
    Write_LogEntry -Message "PrusaSlicer Shortcut BaseName: $($shortcutName)" -Level "DEBUG"
    
    # Remove the version number
    $newShortcutName = $shortcutName -replace '\s\d+(\.\d+)+$', ''
    Write_LogEntry -Message "Neuer Shortcut-Name berechnet: $($newShortcutName)" -Level "DEBUG"
    
    # Construct the new shortcut path
    #$newShortcutPath = Join-Path -Path $shortcutDirectory -ChildPath "$newShortcutName.lnk"
    
    # Rename the PrusaSlicer shortcut, preserving the .lnk extension
    Rename-Item -Path $prusaSlicerShortcut.FullName -NewName "$newShortcutName.lnk"
    Write_LogEntry -Message "PrusaSlicer Shortcut umbenannt: $($prusaSlicerShortcut.FullName) -> $($newShortcutName).lnk" -Level "INFO"
}

if ($InstallationFlag -eq $true) {
	Write-Host "	Prusa Dateizuordnung wird gesetzt" -foregroundcolor "Yellow"
    Write_LogEntry -Message "Setze Prusa Dateizuordnungen (InstallationFlag true)." -Level "INFO"

	$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
	Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"
	if (Test-Path $SetUserFTA) {
		$fileExtensions = @(".stl", ".gcode")
		Write_LogEntry -Message ('Dateierweiterungen für Prusa: ' + ($fileExtensions -join ', ')) -Level "DEBUG"
		foreach ($extension in $fileExtensions) {
			Write_LogEntry -Message "Registriere Dateityp $($extension) mit PrusaSlicer." -Level "INFO"
			& $SetUserFTA --reg "C:\Program Files\Prusa3D\PrusaSlicer\prusa-slicer.exe" $extension
			Write_LogEntry -Message "SetUserFTA ausgeführt für $($extension)" -Level "DEBUG"
		}
	} else {
        Write_LogEntry -Message "SetUserFTA nicht gefunden: $($SetUserFTA)" -Level "WARNING"
    }

	$prusaGCodeRegFile = "$Serverip\Daten\Prog\3D\PrusaSlicer_gcode.reg"
	$prusaStlRegFile = "$Serverip\Daten\Prog\3D\PrusaSlicer_stl.reg"
	Write_LogEntry -Message "Registry-Dateien: GCode=$($prusaGCodeRegFile); STL=$($prusaStlRegFile)" -Level "DEBUG"

	#if (Test-Path $prusaGCodeRegFile) {
		#Start-Process -FilePath "regedit.exe" -ArgumentList "/s", $prusaGCodeRegFile -Wait
	#}

	#if (Test-Path $prusaStlRegFile) {
		#Start-Process -FilePath "regedit.exe" -ArgumentList "/s", $prusaStlRegFile -Wait
	#}
		
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
		-Path $prusaGCodeRegFile
    Write_LogEntry -Message "RegistryImport aufgerufen für: $($prusaGCodeRegFile)" -Level "INFO"
		
	& $PSHostPath `
		-NoLogo -NoProfile -ExecutionPolicy Bypass `
		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
		-Path $prusaStlRegFile
    Write_LogEntry -Message "RegistryImport aufgerufen für: $($prusaStlRegFile)" -Level "INFO"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
