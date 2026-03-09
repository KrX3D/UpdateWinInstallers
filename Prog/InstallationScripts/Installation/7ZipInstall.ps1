param(
    [switch]$InstallationFlag
)

$ProgramName = "7-Zip"
$ScriptType = "Install"

$parentPath  = Split-Path -Path $PSScriptRoot -Parent
$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"

$uninstallPath = "C:\Program Files\7-Zip\Uninstall.exe"
if (Test-Path $uninstallPath) {
    Write-Host "7-Zip ist installiert, Deinstallation beginnt." -ForegroundColor Magenta
    [void](Invoke-ProgramUninstall -UninstallFile $uninstallPath -Arguments "/S" -Wait)
    Write-Host "	7-Zip wurde deinstalliert." -ForegroundColor Green
    Start-Sleep -Seconds 3
}

$installedFiles = Invoke-ProgramInstallFromPattern -PathPattern "$Serverip\Daten\Prog\7z*.exe" -Arguments "/S" -Wait
if (-not $installedFiles -or $installedFiles.Count -eq 0) {
    Write_LogEntry -Message "Keine Installer gefunden unter: $($Serverip)\Daten\Prog\7z*.exe" -Level "WARNING"
} else {
    foreach ($file in $installedFiles) {
        Write-Host "7ZIP wird installiert: $file" -ForegroundColor Magenta
        Write_LogEntry -Message "Installiert: $file" -Level "SUCCESS"
    }
}

Remove-StartMenuEntries -Paths @("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\7-Zip")

if ($InstallationFlag) {
    $setUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
    $fileTypes = @(
        ".001", ".7Z", ".zip", ".rar", ".cab", ".iso", ".img", ".xz", ".txz",
        ".lzma", ".tar", ".cpio", ".bz2", ".bzip2", ".tbz2", ".tbz", ".gz",
        ".gzip", ".tgz", ".tpz", ".z", ".taz", ".lzh", ".lha", ".rpm", ".deb",
        ".arj", ".vhd", ".vhdx", ".wim", ".swm", ".esd", ".fat", ".ntfs",
        ".dmg", ".hfs", ".xar", ".sqashfs", ".apfs"
    )

    if (Set-UserFileAssociations -SetUserFtaPath $setUserFTA -ApplicationPath "C:\Program Files\7-Zip\7zFM.exe" -Extensions $fileTypes) {
        Write_LogEntry -Message "Dateizuordnungen erfolgreich gesetzt." -Level "SUCCESS"
    } else {
        Write_LogEntry -Message "SFTA.exe nicht gefunden: $setUserFTA" -Level "WARNING"
    }
}
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"