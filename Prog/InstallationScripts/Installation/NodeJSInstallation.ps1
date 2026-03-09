param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Node.js"
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

# Find current Node.js installations
Write_LogEntry -Message "Suche nach vorhandenen Node.js-Installationen (Win32_Product) ..." -Level "DEBUG"
try {
    $nodeInstallInfo = Get-CimInstance -ClassName Win32_Product -ErrorAction Stop | Where-Object { $_.Name -like "Node.js*" }
    $nodeCount = @($nodeInstallInfo).Count
    Write_LogEntry -Message "Gefundene Node.js-Installationen: $($nodeCount)" -Level "DEBUG"
} catch {
    Write_LogEntry -Message "Fehler bei Abfrage vorhandener Node.js Installationen: $($_)" -Level "ERROR"
    $nodeInstallInfo = $null
}

# Uninstall existing Node.js if found
if ($nodeInstallInfo) {
    Write_LogEntry -Message "Node.js ist installiert, Deinstallation beginnt." -Level "INFO"
    Write-Host "Node.js ist installiert, Deinstallation beginnt." -ForegroundColor "Magenta"
    
    foreach ($installation in $nodeInstallInfo) {
        $uninstallId = $installation.IdentifyingNumber
        $uninstallArgs = "/x $uninstallId /qn"
        Write_LogEntry -Message "Starte Deinstallation für $($installation.Name) Version $($installation.Version) mit msiexec Args: $($uninstallArgs)" -Level "INFO"
        Write-Host "    Deinstalliere: $($installation.Name) ($($installation.Version))" -ForegroundColor "Cyan"
        try {
            [void](Invoke-InstallerFile -FilePath "msiexec.exe" -Arguments $uninstallArgs -Wait)
            Write_LogEntry -Message "Deinstallation erfolgreich für $($installation.Name) (ProductId: $($uninstallId))." -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler bei Deinstallation von $($installation.Name) (ProductId: $($uninstallId)): $($_)" -Level "ERROR"
        }
    }
    
    Write_LogEntry -Message "Node.js Deinstallation(en) abgeschlossen. Warte 3 Sekunden." -Level "DEBUG"
    Write-Host "    Node.js wurde deinstalliert." -ForegroundColor "Green"
    Start-Sleep -Seconds 3
}

# Installation
Write_LogEntry -Message "Suche nach Node.js Installer im Pfad: $($InstallationFolder)\node-v*-x64.msi" -Level "DEBUG"
$installer = Get-ChildItem -Path "$InstallationFolder\node-v*-x64.msi" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName

if ($installer) {
    Write_LogEntry -Message "Gefundener Installer: $($installer)" -Level "INFO"
    Write-Host "Node.js wird installiert: $installer" -ForegroundColor "Magenta"
    
    # Install Node.js with standard options (silent install)
    #$features    = "NodeRuntime,corepack,npm,EnvironmentPathNode,EnvironmentPathNpmModules" #corepack ist in v25 nicht mehr enthalten. Error 2711 bei Installation
    $features    = "NodeRuntime,npm,EnvironmentPathNode,EnvironmentPathNpmModules"
    $removeFeats = "DocumentationShortcuts"
    $installArgs = "/i `"$installer`" /qb /norestart ADDLOCAL=$features REMOVE=$removeFeats"
    Write_LogEntry -Message "Starte msiexec mit Argumenten: $($installArgs)" -Level "INFO"
    try {
        [void](Invoke-InstallerFile -FilePath "msiexec.exe" -Arguments $installArgs -Wait)
        Write_LogEntry -Message "Node.js Installer ausgeführt: $($installer)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Ausführen des Node.js Installers $($installer): $($_)" -Level "ERROR"
        Write-Host "    Fehler beim Installieren von Node.js." -ForegroundColor "Red"
        exit 1
    }
    
    Write-Host "    Node.js wurde installiert." -ForegroundColor "Green"
} else {
    Write_LogEntry -Message "Keine Node.js Installationsdatei gefunden im Muster: $($InstallationFolder)\node-v*-x64.msi" -Level "ERROR"
    Write-Host "Keine Node.js Installationsdatei gefunden." -ForegroundColor "Red"
    exit
}

$startMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
Write_LogEntry -Message "Prüfe Startmenüpfad auf Node.js Verknüpfungen: $($startMenuPath)" -Level "DEBUG"

# Alle Einträge mit "Node.js*" im Namen entfernen
try {
    $nodeLinks = Get-ChildItem -Path $startMenuPath -Filter "Node.js*" -ErrorAction SilentlyContinue
    $linkCount = @($nodeLinks).Count
    if ($nodeLinks) {
        Write_LogEntry -Message "Gefundene Startmenüeinträge zum Entfernen: $($linkCount)" -Level "INFO"
        Write-Host "	Startmenüeinträge werden entfernt..." -ForegroundColor Cyan
        $nodeLinks | ForEach-Object {
            try {
                Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop
                Write_LogEntry -Message "Entfernt: $($_.FullName)" -Level "SUCCESS"
            } catch {
                Write_LogEntry -Message "Fehler beim Entfernen von $($_.FullName): $($_)" -Level "ERROR"
            }
        }
    } else {
        Write_LogEntry -Message "Keine Node.js Startmenüeinträge gefunden im Pfad: $($startMenuPath)" -Level "DEBUG"
    }
} catch {
    Write_LogEntry -Message "Fehler beim Auflisten/Entfernen der Startmenüeinträge: $($_)" -Level "ERROR"
}

# Verify installation
Write_LogEntry -Message "Überprüfe Node.js und npm Installation (starte node -v und npm -v)" -Level "INFO"
try {
    # Environment Pfade neu einlesen, da ansonsten npm und node Befehle nicht funktionieren
    $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $userPath    = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machinePath;$userPath"
    Write_LogEntry -Message "Environment PATH neu gesetzt (Machine + User) für aktuelle Session." -Level "DEBUG"

    $nodeVersion = & node -v
    $npmVersion = & npm -v

    Write_LogEntry -Message "Node.js Version ermittelt: $($nodeVersion)" -Level "SUCCESS"
    Write_LogEntry -Message "npm Version ermittelt: $($npmVersion)" -Level "SUCCESS"

    Write-Host ""
    Write-Host "Node.js Installation erfolgreich:" -ForegroundColor "Green"
    Write-Host "    Node.js Version: $nodeVersion" -ForegroundColor "Cyan"
    Write-Host "    npm Version: $npmVersion" -ForegroundColor "Cyan"
} catch {
    Write_LogEntry -Message "Fehler bei der Überprüfung der Node.js Installation: $($_)" -Level "ERROR"
    Write-Host "Fehler bei der Überprüfung der Node.js Installation: $_" -ForegroundColor "Red"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
