param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "MQTT Explorer"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

Write_LogEntry -Message "Beginne Installation: MQTT Explorer" -Level "INFO"
Write-Host "MQTT Explorer wird installiert" -foregroundcolor "magenta"

$MQTTExplorerExeFiles = Get-ChildItem "$Serverip\Daten\Prog\MQTT-Explorer*.exe" -ErrorAction SilentlyContinue
if ($MQTTExplorerExeFiles) {
    $count = @($MQTTExplorerExeFiles).Count
    Write_LogEntry -Message "Gefundene MQTT Explorer Installer: $($count)" -Level "DEBUG"
    foreach ($file in $MQTTExplorerExeFiles) {
        Write_LogEntry -Message "Starte Installer: $($file.FullName) mit Argument '/S' (synchronous - Wait)" -Level "INFO"
        try {
            [void](Invoke-InstallerFile -FilePath $file.FullName -Arguments "/S" -Wait)
            Write_LogEntry -Message "Installer erfolgreich ausgeführt: $($file.FullName)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Ausführen des Installers $($file.FullName): $($_)" -Level "ERROR"
        }
    }
} else {
    Write_LogEntry -Message "Keine MQTT Explorer Installer gefunden unter: $($Serverip)\Daten\Prog (Pattern MQTT-Explorer*.exe)" -Level "WARNING"
}

$MQTTExplorerLnkFile = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\MQTT Explorer.lnk"
Write_LogEntry -Message "Prüfe Startmenüverknüpfung: $($MQTTExplorerLnkFile)" -Level "DEBUG"
if (Test-Path $MQTTExplorerLnkFile) {
    try {
        Write_LogEntry -Message "Startmenüeintrag gefunden: $($MQTTExplorerLnkFile). Entferne..." -Level "INFO"
        Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
        Remove-Item -Path $MQTTExplorerLnkFile -Force
        Write_LogEntry -Message "Startmenüeintrag entfernt: $($MQTTExplorerLnkFile)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Entfernen der Startmenüverknüpfung $($MQTTExplorerLnkFile): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Keine Startmenüverknüpfung gefunden: $($MQTTExplorerLnkFile)" -Level "DEBUG"
}

$MQTTExplorerConfig = "$Serverip\Daten\Prog\MQTT_Explorer\settings.json"
Write_LogEntry -Message "Prüfe Config-Datei: $($MQTTExplorerConfig)" -Level "DEBUG"
if (Test-Path $MQTTExplorerConfig) {
    Write_LogEntry -Message "Config gefunden: $($MQTTExplorerConfig). Setze Konfiguration ..." -Level "INFO"
    $MQTTExplorerDestination = "$env:USERPROFILE\AppData\Roaming\MQTT-Explorer\"
    Write_LogEntry -Message "MQTT Explorer Zielpfad: $($MQTTExplorerDestination)" -Level "DEBUG"

    if (!(Test-Path $MQTTExplorerDestination)) {
        try {
            New-Item -Path $MQTTExplorerDestination -ItemType Directory -Force | Out-Null
            Write_LogEntry -Message "Zielverzeichnis erstellt: $($MQTTExplorerDestination)" -Level "DEBUG"
        } catch {
            Write_LogEntry -Message "Fehler beim Erstellen des Zielverzeichnisses $($MQTTExplorerDestination): $($_)" -Level "ERROR"
        }
    }

    try {
        Copy-Item -Path $MQTTExplorerConfig -Destination $MQTTExplorerDestination -Force
        Write_LogEntry -Message "Config kopiert von $($MQTTExplorerConfig) nach $($MQTTExplorerDestination)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Kopieren der Config $($MQTTExplorerConfig) nach $($MQTTExplorerDestination): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Keine Config-Datei gefunden: $($MQTTExplorerConfig). Überspringe Konfiguration." -Level "WARNING"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
