param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Microsoft Edge Webview 2"
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

#Bei Update wird ohne deinstallation die neue Version richtig installiert

# Define the path to the EXE file directory
Write_LogEntry -Message "Suche nach Microsoft Edge Webview2 Installer im Pfad: $($ServerIP)\Daten\Prog\ImageGlass\" -Level "DEBUG"
$exeFilePath = Get-ChildItem -Path "$ServerIP\Daten\Prog\ImageGlass\" -Filter "MicrosoftEdgeWebview2Setup.exe" |
    Sort-Object -Property LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName

if ($exeFilePath) {
    Write_LogEntry -Message "Gefundene Installationsdatei für $($ProgramName): $($exeFilePath)" -Level "INFO"
    Write-Host "$ProgramName wird installiert..." -ForegroundColor Magenta
    try {
        Write_LogEntry -Message "Starte Installation von $($ProgramName) mit Datei: $($exeFilePath)" -Level "INFO"
        # Install the program silently
        #Start-Process -FilePath $exeFilePath -ArgumentList "/install", "/silent" -Wait
        Start-Process -FilePath $exeFilePath -Wait
        Write_LogEntry -Message "Start-Process für $($exeFilePath) wurde ausgeführt." -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler bei der Installation von $($ProgramName): $($_)" -Level "ERROR"
        Write-Host "Fehler bei der Installation von $ProgramName : $_" -ForegroundColor Red
    }
} else {
    Write_LogEntry -Message "Keine Installationsdatei für $($ProgramName) gefunden im Verzeichnis: $($ServerIP)\Daten\Prog\ImageGlass" -Level "ERROR"
    Write-Host "Keine Installationsdatei für $ProgramName gefunden im Verzeichnis: $ServerIP\Daten\Prog\ImageGlass" -ForegroundColor Red
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
