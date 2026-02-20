param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Git"
$ScriptType  = "Install"

$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath


Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt: $($ProgramName); ScriptType gesetzt: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

Write_LogEntry -Message "Starte Git-Installation (falls vorhanden) mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write-Host "Git wird installiert" -foregroundcolor "magenta"

# Install Git
$gitInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\Git*.exe" | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue
if ($gitInstaller) {
    Write_LogEntry -Message "Git-Installer gefunden: $($gitInstaller)" -Level "DEBUG"
    Write_LogEntry -Message "Starte Git-Installer: $($gitInstaller) mit stillen Parametern (Wait)" -Level "INFO"
    [void](Invoke-InstallerFile -FilePath $gitInstaller -Arguments '/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXS', '/NOCANCEL', '/NORESTART', '/NOICONS', '/PathOption=CmdTools' -Wait)
    Write_LogEntry -Message "Git-Installer beendet: $($gitInstaller)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Git-Installer gefunden unter: $($Serverip)\Daten\Prog (Pattern: Git*.exe)" -Level "WARNING"
    Write-Host "Kein Git-Installer gefunden: $Serverip\Daten\Prog\Git*.exe" -ForegroundColor "Yellow"
}

#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Git_Kontextmenu_entfernen.reg" -Wait
Write_LogEntry -Message "Starte Registry-Import Skript fr Git Kontextmenu (via PSHostPath): $($PSHostPath) mit Pfad: $($Serverip)\Daten\Customize_Windows\Reg\Kontextmenu\Git_Kontextmenu_entfernen.reg" -Level "INFO"
try {
    & $PSHostPath `
        -NoLogo -NoProfile -ExecutionPolicy Bypass `
        -File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
        -Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Git_Kontextmenu_entfernen.reg"
    Write_LogEntry -Message "Registry-Importskript aufgerufen: $($Serverip)\Daten\Customize_Windows\Scripte\RegistryImport.ps1" -Level "SUCCESS"
} catch {
    Write_LogEntry -Message "Fehler beim Aufruf des Registry-Importskripts: $($_)" -Level "ERROR"
    Write-Host "Fehler beim Ausfhren des Registry-Importskripts: $($_)" -ForegroundColor "Red"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
