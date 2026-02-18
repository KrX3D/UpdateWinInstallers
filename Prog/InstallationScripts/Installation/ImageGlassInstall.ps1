param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "ImageGlass"
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

#Bei Update wird ohne deinstallation die neue Version richtig installiert

$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
Write_LogEntry -Message "SetUserFTA Pfad: $($SetUserFTA)" -Level "DEBUG"

Write-Host "Image Glass wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Beginne Installation von ImageGlass" -Level "INFO"

$ImageGlassMsi = Get-ChildItem "$Serverip\Daten\Prog\ImageGlass*.msi" | Select-Object -ExpandProperty FullName
Write_LogEntry -Message "Gefundene ImageGlass MSI-Datei: $($ImageGlassMsi)" -Level "DEBUG"

Write_LogEntry -Message "Starte MSI-Installation via msiexec für: $($ImageGlassMsi)" -Level "INFO"
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$ImageGlassMsi`" /qn /quiet /norestart" -Wait
Write_LogEntry -Message "msiexec Prozess beendet für: $($ImageGlassMsi)" -Level "SUCCESS"

$publicDesktopPattern = Join-Path $env:PUBLIC "Desktop\ImageGlass*.lnk"
if (Test-Path $publicDesktopPattern) {
    Write_LogEntry -Message ('Public Desktop Shortcut(s) gefunden: ' + $publicDesktopPattern + '. Entferne...') -Level "INFO"
	Write-Host "	Desktopeintrag PUBLIC wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $publicDesktopPattern -Force
    Write_LogEntry -Message "Public Desktop Shortcut(s) entfernt." -Level "SUCCESS"
} else {
    Write_LogEntry -Message ('Keine Public Desktop Shortcuts gefunden: ' + $publicDesktopPattern) -Level "DEBUG"
}

$userDesktopPattern = Join-Path $env:USERPROFILE "Desktop\ImageGlass*.lnk"
if (Test-Path $userDesktopPattern) {
    Write_LogEntry -Message ('User Desktop Shortcut(s) gefunden: ' + $userDesktopPattern + '. Entferne...') -Level "INFO"
	Write-Host "	Desktopeintrag USERPROFILE wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $userDesktopPattern -Force
    Write_LogEntry -Message "User Desktop Shortcut(s) entfernt." -Level "SUCCESS"
} else {
    Write_LogEntry -Message ('Keine User Desktop Shortcuts gefunden: ' + $userDesktopPattern) -Level "DEBUG"
}

$startMenuProgramsPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ImageGlass"
if (Test-Path $startMenuProgramsPath) {
    Write_LogEntry -Message "Start Menu Programs Ordner für ImageGlass gefunden. Entferne Ordner." -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $startMenuProgramsPath -Recurse -Force
    Write_LogEntry -Message ('Start Menu Programs Ordner entfernt: ' + $startMenuProgramsPath) -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Start Menu Programs Ordner für ImageGlass gefunden." -Level "DEBUG"
}

$userStartMenuProgramsPath = Join-Path $env:USERPROFILE "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\ImageGlass"
if (Test-Path $userStartMenuProgramsPath) {
    Write_LogEntry -Message "User Start Menu Programs Ordner für ImageGlass gefunden. Entferne Ordner." -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $userStartMenuProgramsPath -Recurse -Force
    Write_LogEntry -Message ('User Start Menu Programs Ordner entfernt: ' + $env:USERPROFILE + '\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\ImageGlass') -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein User Start Menu Programs Ordner für ImageGlass gefunden." -Level "DEBUG"
}

$shortcut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ImageGlass.lnk"
Write_LogEntry -Message "Prüfe Startmenüverknüpfung (Leaf): $($shortcut)" -Level "DEBUG"
if (Test-Path -Path $shortcut -PathType Leaf) {
    Write_LogEntry -Message "Startmenüverknüpfung gefunden: $($shortcut). Entferne." -Level "INFO"
	Write-Host "	Startmenüverknüpfung 'ImageGlass' wird entfernt." -ForegroundColor Cyan
	Remove-Item -Path $shortcut -Force
    Write_LogEntry -Message "Startmenüverknüpfung entfernt: $($shortcut)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Keine Startmenüverknüpfung gefunden: $($shortcut)" -Level "DEBUG"
}

$extensions = @(
    ".3fr",
    ".ari",
    ".arw",
    ".avci",
    ".avif",
    ".b64",
    ".bay",
    ".bmp",
    ".cap",
    ".cr2",
    ".crw",
    ".cur",
    ".cut",
    ".dcr",
    ".dcs",
    ".dds",
    ".dib",
    ".dng",
    ".drf",
    ".eip",
    ".emf",
    ".erf",
    ".exif",
    ".exr",
    ".gif",
    ".gpr",
    ".hdr",
    ".heic",
    ".heif",
    ".ico",
    ".iiq",
    ".jfif",
    ".jpe",
    ".jpeg",
    ".jpg",
    ".jxr",
    ".k25",
    ".kdc",
    ".mdc",
    ".mef",
    ".mos",
    ".mrw",
    ".nef",
    ".nrw",
    ".obm",
    ".orf",
    ".pbm",
    ".pcx",
    ".pef",
    ".pgm",
    ".png",
    ".ppm",
    ".psb",
    ".psd",
    ".ptx",
    ".pxn",
    ".r3d",
    ".raf",
    ".raw",
    ".rle",
    ".rw2",
    ".rwl",
    ".rwz",
    ".sr2",
    ".srf",
    ".srw",
    ".svg",
    ".tga",
    ".tif",
    ".tiff",
    ".wdp",
    ".webp",
    ".wmf",
    ".wpg",
    ".x3f",
    ".xbm",
    ".xpm"
)
Write_LogEntry -Message "Anzahl der zu registrierenden Erweiterungen: $($extensions.Count)" -Level "DEBUG"

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag=true: Beginne Registrierung der Dateiendungen" -Level "INFO"
	Write-Host "	Dateiendung werden registriert." -foregroundcolor "Yellow"
	
	if (Test-Path "C:\Program Files\ImageGlass\ImageGlass.exe") {
		$InstallPath = "C:\Program Files\ImageGlass\ImageGlass.exe"
        Write_LogEntry -Message "InstallPath gesetzt auf Program Files Pfad: $($InstallPath)" -Level "DEBUG"
	}
	else
	{
		$InstallPath = "$env:USERPROFILE\AppData\Local\Programs\ImageGlass\ImageGlass.exe"
        Write_LogEntry -Message "InstallPath gesetzt auf User Local Pfad: $($InstallPath)" -Level "DEBUG"
	}
	
	foreach ($extension in $extensions) {
        Write_LogEntry -Message "Setze FTA für Extension $($extension) mit Programm $($InstallPath)" -Level "DEBUG"
		& $SetUserFTA --reg $InstallPath $extension
        Write_LogEntry -Message "FTA gesetzt für Extension $($extension)" -Level "SUCCESS"
	}

	#igtasks.exe regassociations "*.bmp;*.cur;*.cut;*.dib;*.emf;*.exif;*.gif;*.ico;*.jfif;*.jpe;*.jpeg;*.jpg;*.pbm;*.pcx;*.pgm;*.png;*.ppm;*.psb;*.svg;*.tif;*.tiff;*.webp;*.wmf;*.wpg;*.xbm;*.xpm;*.exr;*.hdr;*.psd;*.tga;"

	#Write-Host "	Wird als Standardapp gesetzt." -foregroundcolor "Yellow"
	#Start-Process -FilePath "$Serverip\Daten\Prog\AutoIt_Scripts\SettingsApp.exe" -Wait

	#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\ImageGlass_Kontextmenu_open.reg" -Wait
	
#	& $PSHostPath `
#		-NoLogo -NoProfile -ExecutionPolicy Bypass `
#		-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
#		-Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\ImageGlass_Kontextmenu_open.reg"
} else {
    Write_LogEntry -Message "InstallationFlag != true: Überspringe Registrierung der Dateiendungen" -Level "INFO"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
