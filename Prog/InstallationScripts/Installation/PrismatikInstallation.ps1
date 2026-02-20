param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Prismatik"
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

$Prismatik = "$Serverip\Daten\Prog\Prismatik"
$destinationFolder = "$env:USERPROFILE\Prismatik"

Write_LogEntry -Message "Prismatik Quellordner: $($Prismatik); Zielordner: $($destinationFolder)" -Level "DEBUG"

Write-Host "Prismatik wird installiert" -foregroundcolor "magenta"
Write_LogEntry -Message "Prismatik Installationsroutine gestartet." -Level "INFO"

# Install Prismatik Unofficial if it exists
$prismatikInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\Prismatik.unofficial.64bit.Setup*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
Write_LogEntry -Message ("Gefundener Prismatik Installer: " + $($(if ($prismatikInstaller) { $prismatikInstaller.FullName } else { '<none>' }))) -Level "DEBUG"

if ($prismatikInstaller) {
    try {
        Write_LogEntry -Message "Starte Prismatik Installer: $($prismatikInstaller.FullName)" -Level "INFO"
        [void](Invoke-InstallerFile -FilePath $prismatikInstaller.FullName -Arguments '/SP-', '/VERYSILENT', '/SUPPRESSMSGBOXS', '/NOCANCEL', '/NORESTART', '/NOICONS' -Wait)
        Write_LogEntry -Message "Prismatik Installer ausgefhrt: $($prismatikInstaller.FullName)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Ausfhren des Prismatik Installers $($prismatikInstaller.FullName): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Prismatik Installer gefunden unter $($Serverip)\Daten\Prog" -Level "WARNING"
}

if ($InstallationFlag -eq $true) {
    if (Test-Path $Prismatik) {
        Write_LogEntry -Message "Prismatik Konfigurationsordner gefunden: $($Prismatik). Starte Wiederherstellung." -Level "INFO"
        if (!(Test-Path $destinationFolder)) {
            try {
                New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                Write_LogEntry -Message "Zielordner erstellt: $($destinationFolder)" -Level "DEBUG"
            } catch {
                Write_LogEntry -Message "Fehler beim Erstellen des Zielordners $($destinationFolder): $($_)" -Level "ERROR"
            }
        }
        try {
            Get-ChildItem $Prismatik | Copy-Item -Destination $destinationFolder -Recurse -Force
            Write_LogEntry -Message "Prismatik Konfiguration kopiert von $($Prismatik) nach $($destinationFolder)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Kopieren der Prismatik Konfiguration von $($Prismatik) nach $($destinationFolder): $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Prismatik Konfigurationsordner nicht gefunden: $($Prismatik)" -Level "WARNING"
    }

    # Wait for 5 seconds
    Start-Sleep -Seconds 5
    Write_LogEntry -Message "Warte 5 Sekunden nach Installation." -Level "DEBUG"

    # Create Prismatik desktop shortcut
    $shortcutPath = "$env:USERPROFILE\Desktop\Prismatik.lnk"
    $targetPath = 'C:\Program Files\Prismatik\Prismatik.exe'
    $startInPath = 'C:\Program Files\Prismatik'
    $shortcutName = 'Prismatik'

    Write_LogEntry -Message "Erstelle Desktop-Shortcut: $($shortcutPath) -> Target: $($targetPath)" -Level "INFO"

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetPath
        $shortcut.WorkingDirectory = $startInPath
        $shortcut.IconLocation = "$targetPath,0"
        $shortcut.Description = $shortcutName
        $shortcut.Save()
        Write_LogEntry -Message "Shortcut wurde erfolgreich angelegt: $($shortcutPath)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Erstellen des Shortcuts $($shortcutPath): $($_)" -Level "ERROR"
    }

    Write-Host "	Link wurde angelegt unter: $shortcutPath" -foregroundcolor "Cyan"
    Write_LogEntry -Message "Link erstellt auf Desktop: $($shortcutPath)" -Level "DEBUG"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
