param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB config dateien zu kopieren. Damit Konfig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Agent Ransack"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write-DeployLog -Message "Script gestartet mit InstallationFlag: $InstallationFlag" -Level 'INFO'

$config             = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$agentRansackExe    = "C:\Program Files\Mythicsoft\Agent Ransack\AgentRansack.exe"
$startMenuFolder    = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Agent Ransack"
$configPath         = "$env:APPDATA\Mythicsoft\AgentRansack\config\config_v9.xml"

# ── Find installer (MSI preferred, EXE fallback) ───────────────────────────────
# Filename pattern: agentransack*.msi / agentransack*.exe (no _x64_ required)
$msiFilter = "agentransack*.msi"
$exeFilter = "agentransack*.exe"

Write-DeployLog -Message "Suche MSI: $InstallationFolder\$msiFilter" -Level 'DEBUG'
$installerFile = Get-InstallerFilePath -Directory $InstallationFolder -Filter $msiFilter
$installerType = 'MSI'

if ($installerFile) {
    Write-DeployLog -Message "MSI gefunden: $($installerFile.FullName)" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Kein MSI gefunden – suche EXE: $InstallationFolder\$exeFilter" -Level 'DEBUG'
    $installerFile = Get-InstallerFilePath -Directory $InstallationFolder -Filter $exeFilter
    $installerType = 'EXE'

    if ($installerFile) {
        Write-DeployLog -Message "EXE gefunden: $($installerFile.FullName)" -Level 'DEBUG'
    } else {
        Write-DeployLog -Message "Kein EXE gefunden." -Level 'DEBUG'
    }
}

if (-not $installerFile) {
    $msg = "Keine Installationsdatei (MSI oder EXE) gefunden in: $InstallationFolder (MSI-Filter: $msiFilter, EXE-Filter: $exeFilter)"
    Write-DeployLog -Message $msg -Level 'ERROR'
    Write-Host ""
    Write-Host $msg -ForegroundColor Red
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet (Fehler: kein Installer)"
    exit 1
}

Write-DeployLog -Message "Gefundene Installationsdatei [$installerType]: $($installerFile.FullName)" -Level 'INFO'

# ── Install ────────────────────────────────────────────────────────────────────
Write-Host "$ProgramName wird installiert" -ForegroundColor Magenta
Write-DeployLog -Message "Starte Installation: $($installerFile.FullName)" -Level 'INFO'

$installOk = $false
try {
    if ($installerType -eq 'MSI') {
        # MSI: call msiexec directly so silent args are passed correctly
        $proc = Start-Process -FilePath "msiexec.exe" `
            -ArgumentList "/i", "`"$($installerFile.FullName)`"", "/qn", "/norestart" `
            -Wait -PassThru -ErrorAction Stop
        $installOk = ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
        Write-DeployLog -Message "msiexec beendet mit ExitCode: $($proc.ExitCode)" -Level $(if ($installOk) { 'SUCCESS' } else { 'WARNING' })
    } else {
        # EXE: try common silent flags
        $proc = Start-Process -FilePath $installerFile.FullName `
            -ArgumentList "/VERYSILENT", "/NORESTART", "/SP-" `
            -Wait -PassThru -ErrorAction Stop
        $installOk = ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
        Write-DeployLog -Message "EXE-Installer beendet mit ExitCode: $($proc.ExitCode)" -Level $(if ($installOk) { 'SUCCESS' } else { 'WARNING' })
    }
} catch {
    Write-DeployLog -Message "Fehler beim Starten des Installers: $_" -Level 'ERROR'
}

if (-not $installOk) {
    $msg = "Installation fehlgeschlagen (ExitCode: $($proc?.ExitCode)). Abbruch."
    Write-DeployLog -Message $msg -Level 'ERROR'
    Write-Host ""
    Write-Host $msg -ForegroundColor Red
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet (Installationsfehler)"
    exit 1
}

Write-DeployLog -Message "Installation abgeschlossen: $($installerFile.FullName)" -Level 'SUCCESS'

# ── Start menu cleanup ─────────────────────────────────────────────────────────
Write-DeployLog -Message "Prüfe Startmenü-Ordner: $startMenuFolder" -Level 'DEBUG'
if (Test-Path $startMenuFolder) {
    Write-Host "    Startmenüeintrag wird entfernt." -ForegroundColor Cyan
    Remove-StartMenuEntries -Paths @($startMenuFolder) -EmitHostMessages
} else {
    Write-DeployLog -Message "Kein Startmenüeintrag gefunden." -Level 'DEBUG'
}

# ── First-run / license dialog handling ───────────────────────────────────────
#
#   InstallationFlag = $true  → fresh install: launch app, dismiss license dialog,
#                               then kill it so we can patch the config cleanly.
#   InstallationFlag = $false → update path: just close any running instance quietly.
#
Write-DeployLog -Message "Prüfe Agent Ransack Pfad: $agentRansackExe" -Level 'DEBUG'

if (-not (Test-Path $agentRansackExe)) {
    $msg = "Agent Ransack EXE nicht gefunden nach Installation: $agentRansackExe"
    Write-DeployLog -Message $msg -Level 'ERROR'
    Write-Host $msg -ForegroundColor Red
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet (EXE nicht gefunden)"
    exit 1
}

if ($InstallationFlag) {
    # Fresh install: launch and dismiss the first-run license dialog via SendKeys
    Write-DeployLog -Message "InstallationFlag gesetzt – starte App zum Dismissal des Lizenz-Dialogs." -Level 'INFO'

    $process = Start-Process -FilePath $agentRansackExe -PassThru
    Write-DeployLog -Message "Prozess gestartet: Name=$($process.ProcessName), Id=$($process.Id)" -Level 'DEBUG'

    Start-Sleep -Seconds 2

    Add-Type -AssemblyName "System.Windows.Forms"
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Write-DeployLog -Message "SendKeys: {ENTER}" -Level 'DEBUG'

    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("{DOWN}{DOWN}{ENTER}")
    Write-DeployLog -Message "SendKeys: {DOWN}{DOWN}{ENTER}" -Level 'DEBUG'

    Start-Sleep -Seconds 3

    Write-DeployLog -Message "Beende Prozess Id: $($process.Id)" -Level 'INFO'
    Stop-Process -Id $process.Id -ErrorAction SilentlyContinue
    $process.WaitForExit()
    Write-DeployLog -Message "Prozess beendet." -Level 'SUCCESS'

} else {
    # Update path: silently close any running instance so config can be written
    $running = Get-Process -Name "AgentRansack" -ErrorAction SilentlyContinue
    if ($running) {
        Write-DeployLog -Message "Laufende AgentRansack-Instanz gefunden (Id: $($running.Id)) – wird geschlossen." -Level 'INFO'
        $running | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-DeployLog -Message "AgentRansack geschlossen." -Level 'DEBUG'
    } else {
        Write-DeployLog -Message "Keine laufende AgentRansack-Instanz gefunden." -Level 'DEBUG'
    }
}

# ── Patch theme config to Silver (n=5) ────────────────────────────────────────
Write-Host ""
Write-DeployLog -Message "Prüfe Theme-Config: $configPath" -Level 'DEBUG'

if (Test-Path $configPath) {
    try {
        $xml = Get-Content -Path $configPath -Raw -ErrorAction Stop
        $xml = $xml -replace '<UITheme n="\d+"/>', '<UITheme n="2"/>'
        $xml | Set-Content -Path $configPath -Encoding UTF8 -ErrorAction Stop
        Write-Host "    Theme auf Silver gesetzt: $configPath" -ForegroundColor Cyan
        Write-DeployLog -Message "UITheme auf n=5 (Silver) gesetzt: $configPath" -Level 'SUCCESS'
    } catch {
        Write-DeployLog -Message "Fehler beim Patchen der Theme-Config: $_" -Level 'WARNING'
    }
} else {
    Write-DeployLog -Message "Theme-Config nicht gefunden (noch nicht erzeugt?): $configPath" -Level 'WARNING'
    Write-Host "    Theme-Config nicht gefunden: $configPath" -ForegroundColor Yellow
}

#& $PSHostPath `
#	-NoLogo -NoProfile -ExecutionPolicy Bypass `
#	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
#	-Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\AgentRansack.reg"

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
