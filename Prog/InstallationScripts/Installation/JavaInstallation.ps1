param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB conifg dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "Java"
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

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
