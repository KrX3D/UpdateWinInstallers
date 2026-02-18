param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB config dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "KiCad"
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
    Write_LogEntry -Message "Script beendet wegen fehlender Konfigurationsdatei: $($configPath)" -Level "ERROR"
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert
#Bei update von zB V7 auf V8 muss eine deinstallation stattfinden

$installer = Get-ChildItem -Path "$Serverip\Daten\Prog\Kicad" `
                          -Filter "kicad-*-x86_64.exe" `
                          -ErrorAction SilentlyContinue |
             Select-Object -First 1

Write_LogEntry -Message "Suche KiCad Installer unter: $($Serverip)\Daten\Prog\Kicad mit Filter 'kicad-*-x86_64.exe'" -Level "DEBUG"
if (-not $installer) {
    Write_LogEntry -Message "Kein KiCad-Installer gefunden unter $($Serverip)\Daten\Prog\Kicad" -Level "ERROR"
    Write-Host "⛔️  Kein KiCad-Installer gefunden unter $Serverip\Daten\Prog\Kicad" -ForegroundColor Red #Datei muss in UTF8 BOM konvertiert sein für Emoji
    exit 1
} else {
    Write_LogEntry -Message "KiCad Installer gefunden: $($installer.FullName)" -Level "INFO"
}

if ($installer.BaseName -match 'kicad-(\d+)\.(\d+)') {
    $kiCadVersion = "$($matches[1]).$($matches[2])"
    Write_LogEntry -Message "Gefundene KiCad-Version aus Dateiname: $($kiCadVersion)" -Level "INFO"
    Write-Host "→ Gefundene KiCad-Version: $kiCadVersion" -ForegroundColor Magenta
}
else {
    Write_LogEntry -Message "Konnte Version nicht parsen aus Installer-Name: $($installer.Name)" -Level "ERROR"
    Write-Host "⛔️  Konnte Version nicht parsen aus '$($installer.Name)'" -ForegroundColor Red #Datei muss in UTF8 BOM konvertiert sein für Emoji
    exit 1
}

$kiCadConfigFolder = "$Serverip\Daten\Prog\Kicad\$kiCadVersion"
Write_LogEntry -Message "Erwarteter Config-Backup-Pfad: $($kiCadConfigFolder)" -Level "DEBUG"

if (-not (Test-Path $kiCadConfigFolder)) {
    Write_LogEntry -Message "Konfigurations-Backup für KiCad $($kiCadVersion) fehlt: $($kiCadConfigFolder)" -Level "ERROR"
    $msg = @"
#############################################
FEHLER: Konfigurations-Backup für KiCad $($kiCadVersion) fehlt!
Pfad:   $($kiCadConfigFolder)
Das Skript wird abgebrochen.
#############################################
"@

    Start-Job -ScriptBlock {
        param($text,$timeout)
        # Create the popup COM object
        $wshell = New-Object -ComObject WScript.Shell
        # 16 = Exclamation icon, + 4096 = system-modal (always on top) :contentReference[oaicite:0]{index=0}
        $wshell.Popup($text, $timeout, "KiCad Config fehlt", 16 + 4096)
    } -ArgumentList $msg,300 | Out-Null

    Write_LogEntry -Message "Job gestartet, welcher Popup für fehlende Konfig erstellt hat. Timeout 300s" -Level "INFO"
    exit 1
}

Write_LogEntry -Message "Starte Installation von KiCad $($kiCadVersion) mit Installer: $($installer.FullName)" -Level "INFO"
Write-Host ""
Write-Host "KiCad $kiCadVersion wird installiert" -foregroundcolor "magenta"
Start-Process -FilePath $installer.FullName -ArgumentList '/S','/allusers' -Wait
Write_LogEntry -Message "Installationsprozess beendet für: $($installer.FullName)" -Level "INFO"

# Remove KiCad $kiCadVersion Start Menu shortcuts if they exist
$kiCadStartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\KiCad $kiCadVersion"
Write_LogEntry -Message "Prüfe Startmenüpfad: $($kiCadStartMenuPath)" -Level "DEBUG"
if (Test-Path $kiCadStartMenuPath) {
    Write_LogEntry -Message "Startmenüeintrag gefunden: $($kiCadStartMenuPath). Entferne..." -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    try {
        Remove-Item -Path $kiCadStartMenuPath -Recurse -Force
        Write_LogEntry -Message "Startmenüeintrag entfernt: $($kiCadStartMenuPath)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Entfernen des Startmenüeintrags $($kiCadStartMenuPath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden unter: $($kiCadStartMenuPath)" -Level "DEBUG"
}

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag ist gesetzt. Starte Backup-Wiederherstellung für KiCad $($kiCadVersion)" -Level "INFO"
	Write-Host "Backup wird wiederhergestellt." -foregroundcolor "Yellow"

	# Copy KiCad $kiCadVersion files if they exist
	$destRoaming = "$env:APPDATA\kicad\$kiCadVersion"
	$dest3DModels = "C:\Program Files\KiCad\$kiCadVersion\share\kicad\3dmodels"
	$destFootprints = "C:\Program Files\KiCad\$kiCadVersion\share\kicad\footprints"
	$destSymbols = "C:\Program Files\KiCad\$kiCadVersion\share\kicad\symbols"
	$dest3rdParty = "$env:USERPROFILE\Documents\KiCad\$kiCadVersion\3rdparty"

	$srcBase = "$Serverip\Daten\Prog\Kicad\$kiCadVersion"
	$src3DModels = "$Serverip\Daten\Prog\Kicad\3dmodels"
	$srcFootprints = "$Serverip\Daten\Prog\Kicad\footprints"
	$srcSymbols = "$Serverip\Daten\Prog\Kicad\symbols"
	$src3rdParty = "$Serverip\Daten\Prog\Kicad\3rdparty"

    function Copy-IfExists($src, $dst) {
        Write_LogEntry -Message "Copy-IfExists called with Source: $($src) Destination: $($dst)" -Level "DEBUG"
        if (Test-Path $src) {
            Write_LogEntry -Message "Quelle vorhanden: $($src). Kopiere nach: $($dst)" -Level "INFO"
            if (-not (Test-Path $dst)) { New-Item -ItemType Directory -Path $dst -Force | Out-Null; Write_LogEntry -Message "Zielverzeichnis erstellt: $($dst)" -Level "DEBUG" }
            Get-ChildItem $src | Copy-Item -Destination $dst -Recurse -Force
            Write_LogEntry -Message "Kopieren abgeschlossen von $($src) nach $($dst)" -Level "SUCCESS"
        } else {
            Write_LogEntry -Message "Quelle nicht vorhanden, überspringe Kopieren: $($src)" -Level "WARNING"
        }
    }

    Copy-IfExists $srcBase $destRoaming
    Copy-IfExists $src3DModels $dest3DModels
    Copy-IfExists $srcFootprints $destFootprints
    Copy-IfExists $srcSymbols $destSymbols
    Copy-IfExists $src3rdParty $dest3rdParty
} else {
    Write_LogEntry -Message "InstallationFlag nicht gesetzt; keine Backup-Wiederherstellung durchgeführt." -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
