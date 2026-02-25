param(
    [switch]$InstallationFlag
)

$ProgramName = "Adobe Acrobat"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) {
    throw "DeployToolkit nicht gefunden: $dtPath"
}
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath
Write-DeployLog -Message "Script gestartet mit InstallationFlag: $InstallationFlag" -Level 'INFO'

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

$installerPattern = Join-Path $InstallationFolder 'AcroRdrDC*.exe'
$installArguments = '/sPB /rs /l /msi /qn /norestart ALLUSERS=1 EULA_ACCEPT=YES UPDATE_MODE=0 DISABLE_ARM_SERVICE_INSTALL=1 SUPPRESS_APP_LAUNCH=YES DISABLEDESKTOPSHORTCUT=1 DISABLE_PDFMAKER=YES ENABLE_CHROMEEXT=0 DISABLE_CACHE=1'

Write-DeployLog -Message "Suche Installer mit Pattern: $installerPattern" -Level 'INFO'
$installedFiles = Invoke-ProgramInstallFromPattern -PathPattern $installerPattern -Arguments $installArguments -Wait -SelectNewest
if (-not $installedFiles -or $installedFiles.Count -eq 0) {
    Write-Host "$ProgramName Installer nicht gefunden." -ForegroundColor Red
    Write-DeployLog -Message "Kein Installer gefunden oder Installation fehlgeschlagen: $installerPattern" -Level 'ERROR'
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
    exit
}

Write-Host "Adobe Acrobat Reader DC wird installiert" -ForegroundColor Magenta
Write-DeployLog -Message "Installer ausgeführt: $($installedFiles -join ', ')" -Level 'SUCCESS'

$shortcutPaths = @(
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Acrobat Reader.lnk',
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Acrobat.lnk',
    "$env:PUBLIC\Desktop\Adobe Acrobat.lnk"
)
Remove-StartMenuEntries -Paths $shortcutPaths -Context $ProgramName -EmitHostMessages

$registryImportScript = "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1"
$registryFile = "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Mit_Adobe_Acrobat_Reader_oeffnen_entfernen.reg"
Import-RegistryFile -PSHostPath $PSHostPath -RegistryImportScript $registryImportScript -RegFilePath $registryFile -Context $ProgramName | Out-Null

$setUserFtaPath = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
$pdfAppPath = 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe'
Set-UserFileAssociations -SetUserFtaPath $setUserFtaPath -ApplicationPath $pdfAppPath -Extensions @('.pdf') -Context $ProgramName | Out-Null

$removeAutoStartScript = "$Serverip\Daten\Customize_Windows\Scripte\RemoveAutoStartItems.ps1"
if (Test-Path -LiteralPath $removeAutoStartScript) {
    Write-DeployLog -Message "Rufe RemoveAutoStartItems auf: $removeAutoStartScript" -Level 'INFO'
    & 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe' -ExecutionPolicy Bypass -NoLogo -NoProfile -File $removeAutoStartScript
}

Write-DeployLog -Message "Script beendet: Program=$ProgramName, ScriptType=$ScriptType" -Level 'INFO'
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
