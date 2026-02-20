param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Adobe Acrobat"
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

Write_LogEntry -Message "Beginne Adobe Acrobat Reader DC Installation - prüfe Installer auf Server." -Level "INFO"
Write-Host "Adobe Acrobat Reader DC wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Suche Installer unter: $($Serverip)\Daten\Prog\AcroRdrDC*.exe" -Level "DEBUG"
$acrobatExe = Get-ChildItem "$Serverip\Daten\Prog\AcroRdrDC*.exe" | Select-Object -ExpandProperty FullName
Write_LogEntry -Message "Gefundener Acrobat-Installer: $($acrobatExe)" -Level "DEBUG"
[void](Invoke-InstallerFile -FilePath $acrobatExe -Arguments "/sPB /rs /l /msi /qn /norestart ALLUSERS=1 EULA_ACCEPT=YES UPDATE_MODE=0 DISABLE_ARM_SERVICE_INSTALL=1 SUPPRESS_APP_LAUNCH=YES DISABLEDESKTOPSHORTCUT=1 DISABLE_PDFMAKER=YES ENABLE_CHROMEEXT=0 DISABLE_CACHE=1" -Wait) # REMOVE=AcrobatBrowserIntegration,ReaderBrowserIntegration funktioniert nicht ERROR
Write_LogEntry -Message "Start-Process für Acrobat-Installer aufgerufen: $($acrobatExe)" -Level "INFO"
#REMOVE_PREVIOUS=YES
#"/sAll /rs 

$acrobatShortcut1 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Acrobat Reader.lnk"
Write_LogEntry -Message "Prüfe Startmenü-Verknüpfung 1: $($acrobatShortcut1)" -Level "DEBUG"
if (Test-Path $acrobatShortcut1) {
    Write_LogEntry -Message "Startmenüeintrag gefunden und wird entfernt: $($acrobatShortcut1)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $acrobatShortcut1 -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($acrobatShortcut1)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden bei: $($acrobatShortcut1)" -Level "DEBUG"
}

$acrobatShortcut2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Acrobat.lnk"
Write_LogEntry -Message "Prüfe Startmenü-Verknüpfung 2: $($acrobatShortcut2)" -Level "DEBUG"
if (Test-Path $acrobatShortcut2) {
    Write_LogEntry -Message "Startmenüeintrag gefunden und wird entfernt: $($acrobatShortcut2)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $acrobatShortcut2 -Force
    Write_LogEntry -Message "Startmenüeintrag entfernt: $($acrobatShortcut2)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden bei: $($acrobatShortcut2)" -Level "DEBUG"
}

$acrobatShortcut3 = "$env:PUBLIC\Desktop\Adobe Acrobat.lnk"
Write_LogEntry -Message "Prüfe Desktop-Verknüpfung: $($acrobatShortcut3)" -Level "DEBUG"
if (Test-Path $acrobatShortcut3) {
    Write_LogEntry -Message "Desktop Icon gefunden und wird entfernt: $($acrobatShortcut3)" -Level "INFO"
	Write-Host "	Desktop Icon wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $acrobatShortcut3 -Force
    Write_LogEntry -Message "Desktop Icon entfernt: $($acrobatShortcut3)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Desktop Icon gefunden bei: $($acrobatShortcut3)" -Level "DEBUG"
}

#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Mit_Adobe_Acrobat_Reader_oeffnen_entfernen.reg" -Wait

$registryImportScript = "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1"
Write_LogEntry -Message "Rufe RegistryImport Script auf: $($registryImportScript)" -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1" `
	-Path "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Mit_Adobe_Acrobat_Reader_oeffnen_entfernen.reg"
Write_LogEntry -Message "RegistryImport Script aufgerufen: $($registryImportScript)" -Level "INFO"
		
#Assign PDS to Adobe
$SetUserFTA = "$Serverip\Daten\Customize_Windows\Tools\SFTA.exe"
Write_LogEntry -Message "SetUserFTA Pfad gesetzt: $($SetUserFTA)" -Level "DEBUG"

if (Test-Path $SetUserFTA -PathType Leaf) {
    Write_LogEntry -Message "SFTA gefunden: $($SetUserFTA) - beginne Dateizuordnungen" -Level "INFO"
	Write-Host ""
	Write-Host "Adobe Reader Dateizuordnung - PDF" -foregroundcolor "Yellow"
	
	#https://www.adobe.com/devnet-docs/acrobatetk/tools/AdminGuide/pdfviewer.html
	#Affected ProgIDs for various products
	#Reader (Continuous) -> .pdf AcroExch.Document.DC
	#Acrobat (Continuous) -> .pdf Acrobat.Document.DC
	
	#$ProgID = "AcroExch.Document"
	#$ProgID = "AcroExch.Document.DC"
	#$ProgID = "Acrobat.Document.DC"
	$ProgID = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
	Write_LogEntry -Message "ProgID für Dateizuordnung gesetzt: $($ProgID)" -Level "DEBUG"
	$fileTypes = @(
		".pdf"
	)
	$sortedFileTypes = $fileTypes | Sort-Object

    Write_LogEntry -Message "Anzahl Dateitypen für Zuordnung: $($sortedFileTypes.Count)" -Level "DEBUG"
	foreach ($type in $sortedFileTypes) {
        Write_LogEntry -Message "Setze Dateizuordnung für Typ: $($type) mit Programm: $($ProgID)" -Level "DEBUG"
		& $SetUserFTA --reg $ProgID $type
        Write_LogEntry -Message "Dateizuordnung gesetzt für Typ: $($type)" -Level "DEBUG"
	}
    Write_LogEntry -Message "Dateizuordnungen mittels SFTA abgeschlossen." -Level "SUCCESS"
} else {
    Write_LogEntry -Message "SFTA nicht gefunden unter: $($SetUserFTA)" -Level "WARNING"
}

Write_LogEntry -Message "Rufe RemoveAutoStartItems Skript auf." -Level "INFO"
& C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NoProfile -File "$Serverip\Daten\Customize_Windows\Scripte\RemoveAutoStartItems.ps1"
Write_LogEntry -Message "RemoveAutoStartItems Skript aufgerufen." -Level "INFO"
	
#Disable Protection Driver to assign http https and pds
#New-ItemProperty -Path “HKLM:\SYSTEM\CurrentControlSet\Services\UCPD” -Name “Start” -Value 4 -PropertyType DWORD -Force
#Disable-ScheduledTask -TaskName ‘\Microsoft\Windows\AppxDeploymentClient\UCPD velocity’ 

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
