param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Agent Ransack"
$ScriptType  = "Install"

# === Logger-Header: automatisch eingefügt ===
$parentPath  = Split-Path -Path $PSScriptROOT -Parent
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
    Finalize_LogSession
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

$localFileWildcard = "agentransack_x64_*.msi"
Write_LogEntry -Message "Suche Installationsdatei mit Wildcard: $($localFileWildcard) im Ordner: $($InstallationFolder)" -Level "DEBUG"

$installFilePath = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard | Select-Object -First 1 -ExpandProperty FullName -ErrorAction SilentlyContinue

if ($installFilePath) {
    Write_LogEntry -Message "Gefundene Installationsdatei: $($installFilePath)" -Level "INFO"
    Write-Host "Agent Ransack Pro wird installiert" -foregroundcolor "magenta"
    Write_LogEntry -Message "Starte MSI-Installer: $($installFilePath) mit Argumenten: /qn /norestart" -Level "INFO"
    Start-Process -FilePath $installFilePath -ArgumentList "/qn", "/norestart" -Wait
    Write_LogEntry -Message "Installer-Aufruf beendet für: $($installFilePath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Keine Installationsdatei gefunden für Muster $($localFileWildcard) in $($InstallationFolder)" -Level "WARNING"
}

#$exeToReplace = "$Serverip\Daten\Prog\AgentRansackPro\crack\AgentRansack.exe"
#if (Test-Path $exeToReplace) {
    #Write-Host "Agent Ransack exe wird getauscht"
    #Copy-Item $exeToReplace "C:\Program Files\Mythicsoft\Agent Ransack" -Force
#}

#$langFileToReplace = "$Serverip\Daten\Prog\AgentRansackPro\crack\lang-en.xml"
#if (Test-Path $langFileToReplace) {
    #Write-Host "Agent Ransack Sprachdatei wird getauscht"
    #Copy-Item $langFileToReplace "C:\Program Files\Mythicsoft\Agent Ransack\config" -Force
#}

$shortcutFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Agent Ransack"
Write_LogEntry -Message "Prüfe Startmenu-Ordner: $($shortcutFolder)" -Level "DEBUG"
if (Test-Path $shortcutFolder) {
    Write_LogEntry -Message "Startmenüeintrag gefunden und wird entfernt: $($shortcutFolder)" -Level "INFO"
    Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $shortcutFolder -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($shortcutFolder)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden bei: $($shortcutFolder)" -Level "DEBUG"
}

# Define the path to the Agent Ransack executable
$agentRansackPath = "C:\Program Files\Mythicsoft\Agent Ransack\AgentRansack.exe"
Write_LogEntry -Message "Prüfe Agent Ransack Pfad: $($agentRansackPath)" -Level "DEBUG"

# Check if the executable exists
if (Test-Path $agentRansackPath) {
    Write_LogEntry -Message "Agent Ransack ausführbar gefunden: $($agentRansackPath) - Starte Anwendung." -Level "INFO"

    # Start Agent Ransack and capture the process ID
    $process = Start-Process -FilePath $agentRansackPath -PassThru
    Write_LogEntry -Message "Agent Ransack Prozess gestartet: Name=$($process.ProcessName), Id=$($process.Id)" -Level "DEBUG"

    # Wait for the license window to appear
    Start-Sleep -Seconds 2
    Write_LogEntry -Message "Warte kurz auf UI-Elemente nach Start (2s)." -Level "DEBUG"

    # Load the necessary Windows Forms assembly
    Add-Type -AssemblyName "System.Windows.Forms"
    Write_LogEntry -Message "System.Windows.Forms Assembly geladen." -Level "DEBUG"

    # Simulate pressing the Down Arrow key twice, followed by Enter
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Write_LogEntry -Message "SendKeys: {ENTER} gesendet." -Level "DEBUG"
    
    # Wait for the license window to appear
    Start-Sleep -Seconds 2
    Write_LogEntry -Message "Warte erneut (2s) nach erstem SendWait." -Level "DEBUG"
    
    # Simulate pressing the Down Arrow key twice, followed by Enter
    [System.Windows.Forms.SendKeys]::SendWait("{DOWN}{DOWN}{ENTER}")
    Write_LogEntry -Message "SendKeys: {DOWN}{DOWN}{ENTER} gesendet." -Level "DEBUG"

    # Wait for Agent Ransack to fully open
    Start-Sleep -Seconds 3
    Write_LogEntry -Message "Warte auf vollständiges Laden der Anwendung (3s)." -Level "DEBUG"

    # Schließe Agent Ransack durch Beenden des Prozesses mit der PID
    Write_LogEntry -Message "Beende Agent Ransack Prozess mit Id: $($process.Id)" -Level "INFO"
    Stop-Process -Id $process.Id
    Write_LogEntry -Message "Stop-Process aufgerufen für Id: $($process.Id)" -Level "DEBUG"

    # Close Agent Ransack by sending Alt+F4
    #[System.Windows.Forms.SendKeys]::SendWait("%{F4}")
    
    # Wait for the process to exit
    $process.WaitForExit()
    Write_LogEntry -Message "Agent Ransack Prozess wurde beendet und exit geprüft: Id $($process.Id)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Die ausführbare Datei Agent Ransack wurde nicht unter $($agentRansackPath) gefunden" -Level "ERROR"
    Write-Host "Die ausführbare Datei Agent Ransack wurde nicht unter $agentRansackPath gefunden" -ForegroundColor Red
}

#Change Theme to Silver:
# Set the path dynamically based on the current user
$configPath = "$env:APPDATA\Mythicsoft\AgentRansack\config\config_v9.xml"
Write_LogEntry -Message "Prüfe Theme-Config Pfad: $($configPath)" -Level "DEBUG"

Write-Host ""
# Check if the file exists
if (Test-Path $configPath) {
    Write_LogEntry -Message "Theme-Config gefunden: $($configPath) - Lese Datei ein." -Level "INFO"
    # Read the file content
    $xml = Get-Content -Path $configPath -Raw
    Write_LogEntry -Message "Theme-Config eingelesen, Länge: $($xml.Length) Zeichen" -Level "DEBUG"

    # Replace the UITheme value using regex
    $xml = $xml -replace '<UITheme n="\d+"/>', '<UITheme n="5"/>'
    Write_LogEntry -Message "UITheme Wert ersetzt in geladenem XML-String." -Level "DEBUG"

    # Save the modified content back to the file
    $xml | Set-Content -Path $configPath -Encoding UTF8
    Write_LogEntry -Message "Theme-Config gespeichert: $($configPath)" -Level "SUCCESS"

    Write-Host "	Die Datei wurde erfolgreich aktualisiert: $configPath" -foregroundcolor "Cyan"
} else {
    Write_LogEntry -Message "Theme-Config Datei nicht gefunden: $($configPath)" -Level "WARNING"
    Write-Host "	Die Datei wurde nicht gefunden: $configPath" -foregroundcolor "Red"
}

#& $PSHostPath `
#	-NoLogo -NoProfile -ExecutionPolicy Bypass `
#	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
#	-Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\AgentRansack.reg"

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
