param(
    [switch]$InstallationFlag, #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
    [switch]$Autoit,
    [switch]$Scite
)

$ProgramName = "Autoit"
$ScriptType  = "Install"
$parentPath  = Split-Path -Path $PSScriptRoot -Parent

$dtPath = Join-Path $parentPath "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit nicht gefunden: $dtPath" }
Import-Module -Name $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $parentPath

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag), Autoit: $($Autoit), Scite: $($Scite)" -Level "INFO"
Write_LogEntry -Message "ProgramName gesetzt: $($ProgramName); ScriptType gesetzt: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $parentPath -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

$AutoItFolder = "$Serverip\Daten\Prog\AutoIt_Scripts"
Write_LogEntry -Message "AutoIt Folder gesetzt: $($AutoItFolder)" -Level "DEBUG"

if (($InstallationFlag -eq $true) -or ($Autoit -eq $true)) {
	# Install AutoIt
	Write_LogEntry -Message "Installationsbedingung für AutoIt erfüllt: InstallationFlag=$($InstallationFlag), Autoit=$($Autoit)" -Level "INFO"
	Write-Host "AutoIt wird installiert" -foregroundcolor "magenta"
	$AutoItInstallerPath = Get-ChildItem -Path "$AutoItFolder\autoit-v3*.exe" | Select-Object -ExpandProperty FullName
	Write_LogEntry -Message "Gefundener AutoIt-Installer: $($AutoItInstallerPath)" -Level "DEBUG"
	[void](Invoke-InstallerFile -FilePath $AutoItInstallerPath -Arguments "/S" -Wait)
	Write_LogEntry -Message "AutoIt Installer-Aufruf beendet für: $($AutoItInstallerPath)" -Level "SUCCESS"
}

if (($InstallationFlag -eq $true) -or ($Scite -eq $true)) {
	# Install SciTE4
	Write_LogEntry -Message "Installationsbedingung für SciTE erfüllt: InstallationFlag=$($InstallationFlag), Scite=$($Scite)" -Level "INFO"
	Write-Host "SciTE4 wird installiert" -foregroundcolor "magenta"
	$SciTE4InstallerPath = Get-ChildItem -Path "$AutoItFolder\SciTE4*.exe" | Select-Object -ExpandProperty FullName
	Write_LogEntry -Message "Gefundener SciTE4-Installer: $($SciTE4InstallerPath)" -Level "DEBUG"
	[void](Invoke-InstallerFile -FilePath $SciTE4InstallerPath -Arguments "/S" -Wait)
	Write_LogEntry -Message "SciTE4 Installer-Aufruf beendet für: $($SciTE4InstallerPath)" -Level "SUCCESS"
}

if ($InstallationFlag -eq $true) {
	# Install VC_redist_x86
	Write_LogEntry -Message "Installationsbedingung für VC_redist erfüllt: InstallationFlag=$($InstallationFlag)" -Level "INFO"
	Write-Host "VC redist wird installiert" -foregroundcolor "magenta"
	$VCRedistInstallerPath = Get-ChildItem -Path "$AutoItFolder\VC_redist*.exe" | Select-Object -ExpandProperty FullName
	Write_LogEntry -Message "Gefundener VC Redist Installer: $($VCRedistInstallerPath)" -Level "DEBUG"
	[void](Invoke-InstallerFile -FilePath $VCRedistInstallerPath -Arguments "/install", "/passive", "/qn", "/norestart" -Wait)
	Write_LogEntry -Message "VC Redist Installer-Aufruf beendet für: $($VCRedistInstallerPath)" -Level "SUCCESS"
}

function Remove-PathWithRetries {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$Retries = 6,
        [int]$DelaySeconds = 2
    )

    if (-not (Test-Path $Path)) {
        Write_LogEntry -Message "Path does not exist, nothing to remove: $Path" -Level "DEBUG"
        return $true
    }

    for ($i = 0; $i -le $Retries; $i++) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write_LogEntry -Message "Removed path: $Path" -Level "SUCCESS"
            return $true
        } catch {
            $errMsg = $_.Exception.Message
            Write_LogEntry -Message "Attempt $($i+1) to remove $Path failed: $errMsg" -Level "WARNING"
            if ($i -lt $Retries) {
                Start-Sleep -Seconds $DelaySeconds
                continue
            } else {
                # Final attempt failed -> try per-file removal and report locked files
                Write_LogEntry -Message "Final attempt failed. Trying per-file removal and reporting locked files: $Path" -Level "INFO"
                $lockedFiles = @()
                try {
                    # Remove files first
                    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                        try {
                            if ($_.PSIsContainer) { return }
                            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                            Write_LogEntry -Message "Removed child file: $($_.FullName)" -Level "DEBUG"
                        } catch {
                            Write_LogEntry -Message "Could not remove child file (likely locked): $($_.FullName) -> $($_.Exception.Message)" -Level "WARNING"
                            $lockedFiles += $_.FullName
                        }
                    }
                } catch {
                    Write_LogEntry -Message "Error walking children of $Path : $($_.Exception.Message)" -Level "ERROR"
                }

                # Try to remove empty directories now
                try {
                    Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                    Write_LogEntry -Message "Removed directory after per-file cleanup: $Path" -Level "SUCCESS"
                    return $true
                } catch {
                    Write_LogEntry -Message "Could not remove directory $Path after cleanup: $($_.Exception.Message)" -Level "WARNING"
                    if ($lockedFiles.Count -eq 0) {
                        # Try to enumerate remaining children to report
                        try {
                            $remaining = Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName -ErrorAction SilentlyContinue
                            if ($remaining) {
                                $lockedFiles += $remaining
                            }
                        } catch {}
                    }
                    if ($lockedFiles.Count -gt 0) {
                        Write_LogEntry -Message ("Locked or remaining files preventing deletion: `n" + ($lockedFiles -join "`n")) -Level "ERROR"
                    } else {
                        Write_LogEntry -Message "Directory could not be removed but no specific locked files were detected." -Level "ERROR"
                    }
                    return $false
                }
            }
        }
    }
}

# Move AutoIt Window Info shortcut
$moveSource = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\AutoIt v3\AutoIt Window Info (x64).lnk"
$moveDest = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
Write_LogEntry -Message "Prüfe und verschiebe Shortcut wenn vorhanden: Quelle=$($moveSource), Ziel=$($moveDest)" -Level "DEBUG"
If (Test-Path $moveSource) {
	Write_LogEntry -Message "Shortcut gefunden zum Verschieben: $($moveSource)" -Level "INFO"
	Write-Host "	Startmenüeintrag werden verschoben." -foregroundcolor "Cyan"
    try {
        Move-Item -Path $moveSource -Destination $moveDest -Force -ErrorAction Stop
		Write_LogEntry -Message "Shortcut verschoben: $($moveSource) -> $($moveDest)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Konnte Shortcut nicht verschieben: $($moveSource). Fehler: $($_.Exception.Message)" -Level "WARNING"
    }
}

# Remove AutoIt v3 Start Menu shortcuts
$autoItStartMenu = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\AutoIt v3"
Write_LogEntry -Message "Prüfe Vorhandensein Startmenu-Ordner zum Entfernen: $($autoItStartMenu)" -Level "DEBUG"
If (Test-Path $autoItStartMenu) {
	Write_LogEntry -Message "Startmenüeintrag vorhanden und wird entfernt: $($autoItStartMenu)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
	
    # Try removal with retries and fallback per-file attempt (no reboot scheduling)
    $removed = Remove-PathWithRetries -Path $autoItStartMenu -Retries 6 -DelaySeconds 2
    if (-not $removed) {
        Write_LogEntry -Message "Ordner konnte nicht vollständig entfernt werden; gesperrte Dateien wurden protokolliert: $autoItStartMenu" -Level "WARNING"
    } else {
		Write_LogEntry -Message "Startmenüeintrag entfernt: $($autoItStartMenu)" -Level "SUCCESS"
	}
}

$userAutoItStartMenu = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\AutoIt v3"
Write_LogEntry -Message "Prüfe Benutzerspezifischen Startmenu-Ordner: $($userAutoItStartMenu)" -Level "DEBUG"
If (Test-Path $userAutoItStartMenu) {
	Write_LogEntry -Message "Benutzer-Startmenüeintrag vorhanden und wird entfernt: $($userAutoItStartMenu)" -Level "INFO"
	Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
	
    $removedUser = Remove-PathWithRetries -Path $userAutoItStartMenu -Retries 4 -DelaySeconds 1
    if (-not $removedUser) {
        Write_LogEntry -Message "Benutzerordner konnte nicht vollständig entfernt werden; gesperrte Dateien wurden protokolliert: $userAutoItStartMenu" -Level "WARNING"
    } else {
		Write_LogEntry -Message "Benutzer-Startmenüeintrag entfernt: $($userAutoItStartMenu)" -Level "SUCCESS"
    }
}

#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\AutoIt_Kontextmenu_entfernen.reg" -Wait
#Start-Process -FilePath "REGEDIT.EXE" -ArgumentList "/S $Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Autoit_submenu_Kontextmenu.reg" -Wait

$registryImportScript = "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1"
$registryImportPath1 = "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\AutoIt_Kontextmenu_entfernen.reg"
Write_LogEntry -Message "Aufruf RegistryImport Script 1: Script=$($registryImportScript), RegPath=$($registryImportPath1)" -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File $registryImportScript `
	-Path $registryImportPath1
Write_LogEntry -Message "RegistryImport Script 1 aufgerufen: Script=$($registryImportScript)" -Level "SUCCESS"

$registryImportScript = "$Serverip\Daten\Customize_Windows\Scripte\RegistryImport.ps1"
$registryImportPath2 = "$Serverip\Daten\Customize_Windows\Reg\Kontextmenu\Autoit_submenu_Kontextmenu.reg"
Write_LogEntry -Message "Aufruf RegistryImport Script 2: Script=$($registryImportScript), RegPath=$($registryImportPath2)" -Level "INFO"
& $PSHostPath `
	-NoLogo -NoProfile -ExecutionPolicy Bypass `
	-File $registryImportScript `
	-Path $registryImportPath2
Write_LogEntry -Message "RegistryImport Script 2 aufgerufen: Script=$($registryImportScript)" -Level "SUCCESS"

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
