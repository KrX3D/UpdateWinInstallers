param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Microsoft Windows Desktop Runtime"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt auf: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

#Bei Update wird ohne deinstallation die neue Version richtig installiert

# Define the path to the EXE file directory
Write_LogEntry -Message "Suche nach Installer Dateien im Verzeichnis: $($ServerIP)\Daten\Prog\ImageGlass\" -Level "DEBUG"
$exeFilePath = Get-ChildItem -Path "$ServerIP\Daten\Prog\ImageGlass\" -Filter "windowsdesktop-runtime*.exe" |
    Sort-Object -Property LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName

if ($exeFilePath) {
    Write_LogEntry -Message "Gefundene Installationsdatei für $($ProgramName): $($exeFilePath)" -Level "INFO"
    Write-Host "$ProgramName wird installiert..." -ForegroundColor Magenta
    try {
        Write_LogEntry -Message "Starte Installation: Start-Process $($exeFilePath) mit Argumenten '/install','/passive','/norestart' (wartend)" -Level "INFO"
        # Install the program silently
        [void](Invoke-InstallerFile -FilePath $exeFilePath -Arguments "/install", "/passive", "/norestart" -Wait)
        Write_LogEntry -Message "Start-Process ausgeführt für $($exeFilePath)." -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler bei der Installation von $($ProgramName): $($_)" -Level "ERROR"
        Write-Host "Fehler bei der Installation von $ProgramName : $_" -ForegroundColor Red
    }
} else {
    Write_LogEntry -Message "Keine Installationsdatei für $($ProgramName) gefunden im Verzeichnis: $($ServerIP)\Daten\Prog\ImageGlass" -Level "ERROR"
    Write-Host "Keine Installationsdatei für $ProgramName gefunden im Verzeichnis: $ServerIP\Daten\Prog\ImageGlass" -ForegroundColor Red
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
