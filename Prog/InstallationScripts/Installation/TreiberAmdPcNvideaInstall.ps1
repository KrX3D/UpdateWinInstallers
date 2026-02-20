param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "NVIDIA Grafiktreiber"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName initialisiert: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

$InstallationFolder = $NetworkShareDaten + "\Treiber\AMD_PC"
Write_LogEntry -Message "InstallationFolder gesetzt: $($InstallationFolder)" -Level "DEBUG"

Write-Host "$ProgramName wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Beginne Installation von $($ProgramName)" -Level "INFO"

# Define the folder pattern
$folderPattern = "desktop-win10-win11-64bit-international-nsd-dch-whql$"  # Match folders ending with this pattern

# Get the folder matching the pattern
$matchingFolder = Get-ChildItem -Path $InstallationFolder -Directory | Where-Object {
    $_.Name -match $folderPattern
} | Select-Object -First 1  # Get the first match if there are multiple

$folderName = if ($matchingFolder) { $matchingFolder.FullName } else { '<null>' }
Write_LogEntry -Message "MatchingFolder: $folderName" -Level "DEBUG"

$SetupExe = "setup.exe"
$InstallFolder = Join-Path $matchingFolder $SetupExe
Write_LogEntry -Message "Installationsdatei Pfad: $($InstallFolder)" -Level "DEBUG"

try {
    Write_LogEntry -Message "Starte Installer: $($InstallFolder) mit Argumenten '-noreboot -clean -noeula -nofinish -passive'" -Level "INFO"
    [void](Invoke-InstallerFile -FilePath $InstallFolder -Arguments "-noreboot -clean -noeula -nofinish -passive" -Wait)
    Write_LogEntry -Message "Installer-Prozess beendet: $($InstallFolder)" -Level "SUCCESS"
} catch {
    Write_LogEntry -Message "Fehler beim Starten des Installers $($InstallFolder): $($_)" -Level "ERROR"
}

# Search for NVIDIA-related processes
$nvidiaProcesses = Get-Process | Where-Object { $_.ProcessName -like "*nvidia*" }
Write_LogEntry -Message ("Gefundene NVIDIA-Prozesse: " + $((if ($nvidiaProcesses) { $nvidiaProcesses.Count } else { 0 }))) -Level "DEBUG"

if ($nvidiaProcesses) {
    Write-Host "NVIDIA-bezogene Prozesse gefunden. Diese werden beendet..." -ForegroundColor Yellow
    Write_LogEntry -Message "NVIDIA-bezogene Prozesse werden beendet. Count: $($nvidiaProcesses.Count)" -Level "INFO"
    foreach ($process in $nvidiaProcesses) {
        try {
            Write_LogEntry -Message "Versuche Prozess zu beenden: $($process.ProcessName) (PID: $($process.Id))" -Level "INFO"
            Stop-Process -Id $process.Id -Force
            Write-Host "Prozess beendet: $($process.ProcessName) (PID: $($process.Id))" -ForegroundColor Green
            Write_LogEntry -Message "Prozess beendet: $($process.ProcessName) (PID: $($process.Id))" -Level "SUCCESS"
        } catch {
            Write-Host "Fehler beim Beenden des Prozesses: $($process.ProcessName) (PID: $($process.Id)). Fehler: $_" -ForegroundColor Red
            Write_LogEntry -Message "Fehler beim Beenden des Prozesses: $($process.ProcessName) (PID: $($process.Id)). Fehler: $($_)" -Level "ERROR"
        }
    }
}

# Define the source and target paths
$sourcePath = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\NVIDIA Corporation"
$sourceFile = Join-Path -Path $sourcePath -ChildPath "NVIDIA.lnk"
$targetPath = Join-Path -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs" -ChildPath "NVIDIA.lnk"

# Check if the source file exists
if (Test-Path -Path $sourceFile) {
	Write-Host "	Startmenüeintrag wird verschoben." -foregroundcolor "Cyan"
    Move-Item -Path $sourceFile -Destination $targetPath -Force
    Write_LogEntry -Message "Startmenüverknüpfung verschoben von $($sourceFile) nach $($targetPath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Startmenüverknüpfung nicht gefunden: $($sourceFile)" -Level "DEBUG"
}

if (Test-Path -Path $sourcePath -PathType Container) {
	Write-Host "	Startmenüeintrag 'NVIDIA Corporation' wird entfernt." -ForegroundColor Cyan
	Remove-Item -Path $sourcePath -Recurse -Force
    Write_LogEntry -Message "Startmenü-Ordner entfernt: $($sourcePath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Startmenü-Ordner nicht gefunden: $($sourcePath)" -Level "DEBUG"
}

if (Test-Path "$env:PUBLIC\Desktop\NVIDIA.lnk") {
	Write-Host "	Desktop Icon wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "$env:PUBLIC\Desktop\NVIDIA.lnk" -Force
    Write_LogEntry -Message "Desktop-Link entfernt: $($env:PUBLIC)\Desktop\NVIDIA.lnk" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Desktop-Link nicht gefunden: $($env:PUBLIC)\Desktop\NVIDIA.lnk" -Level "DEBUG"
}

if (Test-Path "$env:PUBLIC\Desktop\NVIDIA App.lnk") {
	Write-Host "	Desktop Icon wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path "$env:PUBLIC\Desktop\NVIDIA App.lnk" -Force
    Write_LogEntry -Message "Desktop-Link entfernt: $($env:PUBLIC)\Desktop\NVIDIA App.lnk" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Desktop-Link nicht gefunden: $($env:PUBLIC)\Desktop\NVIDIA App.lnk" -Level "DEBUG"
}

Write_LogEntry -Message ("Rufe RegistryImport Skript auf: $($Serverip)\Daten\Customize_Windows\Scripte\RegistryImport.ps1 mit Pfad $($Serverip)\Daten\Prog\InstallationScripts\Nvidia_NotifyNewDisplayUpdates.reg") -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
	-Path "$Serverip\Daten\Prog\InstallationScripts\Nvidia_NotifyNewDisplayUpdates.reg"
Write_LogEntry -Message "RegistryImport Skript wurde ausgeführt." -Level "SUCCESS"

Write_LogEntry -Message ("NVIDIA Dienste werden gestopt: $($InstallationFolder)\Manage-NvidiaContainers.ps1 mit dem Parameter -Action Stop") -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$InstallationFolder\Manage-NvidiaContainers.ps1" `
	-Action Stop
Write_LogEntry -Message "NVIDIA Dienste wurden gestopt." -Level "SUCCESS"

Write_LogEntry -Message ("Rufe RegistryImport Skript auf: $($Serverip)\Daten\Customize_Windows\Scripte\RegistryImport.ps1 mit Pfad $($Serverip)\Daten\Prog\InstallationScripts\Nvidia_remove_Control_Panel_from_desktop_context_menu_for_current_user.reg") -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
	-Path "$Serverip\Daten\Prog\InstallationScripts\Nvidia_remove_Control_Panel_from_desktop_context_menu_for_current_user.reg"
Write_LogEntry -Message "RegistryImport Skript wurde ausgeführt." -Level "SUCCESS"

Write_LogEntry -Message ("Rufe RegistryImport Skript auf: $($Serverip)\Daten\Customize_Windows\Scripte\RegistryImport.ps1 mit Pfad $($Serverip)\Daten\Prog\InstallationScripts\Nvidia_Overlay_disable.reg") -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
	-Path "$Serverip\Daten\Prog\InstallationScripts\Nvidia_Overlay_disable.reg"
Write_LogEntry -Message "RegistryImport Skript wurde ausgeführt." -Level "SUCCESS"

Write_LogEntry -Message ("NVIDIA Dienste werden gestartet: $($InstallationFolder)\Manage-NvidiaContainers.ps1 mit dem Parameter -Action Start") -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$InstallationFolder\Manage-NvidiaContainers.ps1" `
	-Action Start
Write_LogEntry -Message "NVIDIA Dienste wurden gestartet." -Level "SUCCESS"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
