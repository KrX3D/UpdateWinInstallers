param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Java"
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
    Write_LogEntry -Message "Konfigurationsdatei $($configPath) gefunden und importiert." -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Script beendet wegen fehlender Konfigurationsdatei: $($configPath)" -Level "ERROR"
    exit
}

#Bei Update wird ohne deinstallation die neue Version richtig installiert

# Alle Java-Versionen deinstallieren
# Definieren Sie die Suchanfrage für Java-Produkte
$searchQuery = "name like 'Java%%'"
Write_LogEntry -Message "Suche nach Java-Produkten mit Query: $($searchQuery)" -Level "DEBUG"

# Holen Sie eine Liste von Java-Produkten, die der Suchanfrage entsprechen
#$javaProducts = Get-WmiObject -Class Win32_Product -Filter $searchQuery

# Achtung: Win32_Product triggert bei einigen MSI-Paketen eine Reconfigure/Reinstall-Aktion!
$javaProducts = Get-CimInstance -ClassName Win32_Product -Filter "Name LIKE 'Java%'" 
Write_LogEntry -Message "Abfrage Win32_Product durchgeführt für 'Java%'" -Level "DEBUG"

if ($javaProducts) {
    Write_LogEntry -Message "Java-Produkte zum Deinstallieren gefunden." -Level "INFO"
    foreach ($product in $javaProducts) {
        $productName = $product.Name
        Write_LogEntry -Message "Starte Deinstallation von Java-Produkt: $($productName)" -Level "INFO"
        Write-Host "Deinstalliere $productName..."
        
        # Invoke-CimMethod statt .Uninstall()
        #$uninstallResult = $product.Uninstall()
        $uninstallResult = Invoke-CimMethod -InputObject $product -MethodName Uninstall

        if ($uninstallResult.ReturnValue -eq 0) {
            Write_LogEntry -Message "Deinstallation erfolgreich: $($productName) (ReturnValue: $($uninstallResult.ReturnValue))" -Level "SUCCESS"
            Write-Host "$productName wurde erfolgreich deinstalliert." -ForegroundColor Green
        } else {
            Write_LogEntry -Message "Fehler bei Deinstallation von $($productName). Fehlercode: $($uninstallResult.ReturnValue)" -Level "ERROR"
            Write-Host "Fehler beim Deinstallieren von $productName. Fehlercode: $($uninstallResult.ReturnValue)" -ForegroundColor Red
        }
    }
} else {
    Write_LogEntry -Message "Keine Java-Versionen gefunden, die deinstalliert werden sollen." -Level "INFO"
    Write-Host "Keine Java-Versionen gefunden, die deinstalliert werden sollen." -ForegroundColor DarkGray
}

Write_LogEntry -Message "Beginne Installation von Java (falls vorhanden)." -Level "INFO"
# Install Java for freerout plugin if it exists
$javaInstaller = Get-ChildItem -Path "$Serverip\Daten\Prog\Kicad\jdk*.msi" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($javaInstaller) {
    Write_LogEntry -Message "Java-Installer gefunden: $($javaInstaller.FullName). Starte still install." -Level "INFO"
    [void](Invoke-InstallerFile -FilePath $javaInstaller.FullName -Arguments '/qn', '/passive', '/norestart' -Wait)
    Write_LogEntry -Message "Java-Installer-Prozess beendet für: $($javaInstaller.FullName)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein Java-Installer gefunden unter: $($Serverip + '\Daten\Prog\Kicad\')" -Level "DEBUG"
}

# Remove Java Development Kit (JDK) Start Menu shortcuts if they exist
$jdkStartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Java Development Kit"
Write_LogEntry -Message "Prüfe JDK Start Menu Pfad: $($jdkStartMenuPath)" -Level "DEBUG"
if (Test-Path $jdkStartMenuPath) {
    Write_LogEntry -Message "JDK Start Menu Pfad gefunden: $($jdkStartMenuPath). Entferne." -Level "INFO"
    Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    Remove-Item -Path $jdkStartMenuPath -Recurse -Force
    Write_LogEntry -Message "JDK Start Menu Pfad entfernt: $($jdkStartMenuPath)" -Level "SUCCESS"
} else {
    Write_LogEntry -Message "Kein JDK Start Menu Pfad gefunden: $($jdkStartMenuPath)" -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
