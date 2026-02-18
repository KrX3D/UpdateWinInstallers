param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Microsoft Visual Studio Code"
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
    Write_LogEntry -Message "Konfigurationsdatei gefunden und importiert: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

Write-Host "VS Code wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Beginne VS Code Installation" -Level "INFO"

# Install VS Code
$vsCodeInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\VSCode*.exe" | Select-Object -ExpandProperty FullName
Write_LogEntry -Message "Gefundener VS Code Installer: $($vsCodeInstaller)" -Level "DEBUG"
try {
    Start-Process -FilePath $vsCodeInstaller -ArgumentList '/VERYSILENT', '/NORESTART', '/MERGETASKS=!runcode,desktopicon,fileassoc' -Wait
    Write_LogEntry -Message "VS Code Installer ausgeführt: $($vsCodeInstaller)" -Level "SUCCESS"
} catch {
    Write_LogEntry -Message "Fehler beim Ausführen des VS Code Installers $($vsCodeInstaller): $($_)" -Level "ERROR"
    throw
}

# Remove VS Code Start Menu shortcuts
$vsCodeStartMenu = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code"
Write_LogEntry -Message "Überprüfe Startmenü-Pfad: $($vsCodeStartMenu)" -Level "DEBUG"
if (Test-Path $vsCodeStartMenu) {
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $vsCodeStartMenu -Recurse -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($vsCodeStartMenu)" -Level "INFO"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden: $($vsCodeStartMenu)" -Level "DEBUG"
}

$shortcut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Visual Studio Installer.lnk"
Write_LogEntry -Message "Überprüfe Shortcut: $($shortcut)" -Level "DEBUG"
if (Test-Path -Path $shortcut -PathType Leaf) {
	Write-Host "	Startmenüverknüpfung 'Visual Studio Installer' wird entfernt." -ForegroundColor Cyan
	Remove-Item -Path $shortcut -Force
    Write_LogEntry -Message "Shortcut entfernt: $($shortcut)" -Level "INFO"
} else {
    Write_LogEntry -Message "Kein Shortcut gefunden: $($shortcut)" -Level "DEBUG"
}

# Install VS Code Extensions
$vsCodeExtensions = @(
    "MS-CEINTL.vscode-language-pack-de",
    "ms-vscode-remote.remote-containers",
    "ms-vscode.cpptools",
    "platformio.platformio-ide"
)

$vsCodeExecutable = Join-Path $env:HOMEPATH "AppData\Local\Programs\Microsoft VS Code\bin\code"
Write_LogEntry -Message "VS Code CLI Pfad gesetzt: $($vsCodeExecutable)" -Level "DEBUG"

if ($InstallationFlag -eq $true) {
    Write_LogEntry -Message "InstallationFlag true: beginne Installation der VS Code Extensions" -Level "INFO"
	foreach ($extension in $vsCodeExtensions) {
		Write-Host "	$extension Plugin wird installiert." -foregroundcolor "Yellow"
        Write_LogEntry -Message "Installiere VS Code Extension: $($extension) mit CLI: $($vsCodeExecutable)" -Level "DEBUG"
		Start-Process -FilePath $vsCodeExecutable -ArgumentList "--install-extension", $extension, "--force" -Wait
        Write_LogEntry -Message "Beendet Installation Extension: $($extension)" -Level "SUCCESS"
	}
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===

#code --install-extension MS-CEINTL.vscode-language-pack-de --force
#code --install-extension ms-vscode-remote.remote-containers --force
#code --install-extension ms-vscode.cpptools --force
#code --install-extension platformio.platformio-ide --force	