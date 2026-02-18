param(
    [switch]$InstallationFlag #wird nur bei $true genutzt, um zB config dateien zu kopiere. Damit Konig Dateien NUR bei einer Installation und NICHT bei einem Update kopiert werden.
)

$ProgramName = "jDownloader 2"
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

Write_LogEntry -Message "Beginne Installation/Prüfung für jDownloader 2" -Level "INFO"
Write-Host "JDownloader 2 wird installiert" -foregroundcolor "magenta"

$JDownloaderExeFile = Get-ChildItem "$Serverip\Daten\Prog\JDownloader2*.exe" | Select-Object -First 1
if ($JDownloaderExeFile) {
    Write_LogEntry -Message "JDownloader-Installer gefunden: $($JDownloaderExeFile.FullName)" -Level "INFO"
} else {
    Write_LogEntry -Message "Kein JDownloader-Installer gefunden unter: $($Serverip)\Daten\Prog" -Level "WARNING"
}

if ($JDownloaderExeFile) {
    $installerArgs = '-q -overwrite "-Dfilelinks=dlc,jdc,ccf,rsdf" "-Ddesktoplink=true" "-DNOAUTO =false" "-Dquicklaunch=false" -splash "JDownloader Install"'
    Write_LogEntry -Message "Starte JDownloader Installer: $($JDownloaderExeFile.FullName) mit Argumenten: $($installerArgs)" -Level "INFO"

    Start-Process -FilePath $JDownloaderExeFile.FullName -ArgumentList $installerArgs
    Write_LogEntry -Message "Installer-Prozess gestartet für $($JDownloaderExeFile.FullName)" -Level "DEBUG"

    # Silent installer autostart app. Killing this behaviour using the code below
    #https://community.chocolatey.org/packages/advanced-port-scanner#files
    $t = 0
    Write_LogEntry -Message "Warte auf JDownloader2 Prozess (Timeout 90s)..." -Level "DEBUG"
    DO
    {
        start-Sleep -Milliseconds 1000 #wait 100ms / loop
        $t++ #increase iteration 
    } Until ($null -ne ($p=Get-Process -Name JDownloader2 -ErrorAction SilentlyContinue) -or ($t -gt 90)) #wait until process is found or timeout reached

    if ($p) { #if process is found
        Write_LogEntry -Message "JDownloader2 Prozess gefunden: Id=$($p.Id). Versuche zu beenden." -Level "INFO"
        try {
            $p | Stop-Process -Force
            Write_LogEntry -Message "JDownloader2 Prozess beendet: Id=$($p.Id)" -Level "SUCCESS"
            Write-Host "Beende JDownloader 2 Prozess"
        } catch {
            Write_LogEntry -Message "Fehler beim Beenden des JDownloader2 Prozesses: $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "Timeout erreicht: Kein JDownloader2 Prozess gefunden nach $($t) Sekunden." -Level "WARNING"
        Write-Host "Timeout für JDownloader 2 Prozess" #no process found but timeout reached
    }
}

#Get Version and add it to Registry
$filePath = "C:\Program Files\JDownloader\build.json"
Write_LogEntry -Message "Lese Build-Info aus: $($filePath)" -Level "DEBUG"
if (Test-Path -Path $filePath) {
    try {
        $jsonContent = Get-Content -Raw -Path $filePath | ConvertFrom-Json
        # Extract the value of 'buildTimestamp'
        $buildTimestamp = $jsonContent.buildDate
        if ($buildTimestamp) {
            $localVersion = $buildTimestamp
            Write_LogEntry -Message "Lokale buildTimestamp gelesen: $($localVersion)" -Level "DEBUG"
        } else {
            $localVersion = "0"
            Write_LogEntry -Message "buildTimestamp nicht gefunden in $($filePath). Setze lokale Version auf 0." -Level "WARNING"
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Lesen/Parsen von $($filePath): $($_)" -Level "ERROR"
        $localVersion = "0"
    }
} else {
    Write_LogEntry -Message "Build-JSON Datei nicht gefunden: $($filePath). Setze lokale Version auf 0." -Level "WARNING"
    $localVersion = "0"
}

# Registry update
$Path  = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | Get-ItemProperty | Where-Object { $_.DisplayName -like $ProgramName + '*' }
$RegistryPath = $Path.PSPath
Write_LogEntry -Message "Gefundene Registry-Pfade für $($ProgramName): $($RegistryPath)" -Level "DEBUG"

if ($RegistryPath) {
    try {
        Write_LogEntry -Message "Setze TimestampVersion = $($localVersion) in Registry-Pfad: $($RegistryPath)" -Level "INFO"
        Set-ItemProperty -Path $RegistryPath -Name 'TimestampVersion' -Value $localVersion -Type String
        Write_LogEntry -Message "Registry-Eintrag erfolgreich geschrieben: $($RegistryPath) TimestampVersion=$($localVersion)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Schreiben des Registry-Eintrags $($RegistryPath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Registry-Pfad gefunden für ProgramName-Muster: $($ProgramName + '*')" -Level "WARNING"
}

if ($InstallationFlag -eq $true) {
    if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
        Write_LogEntry -Message "7-Zip gefunden. Stelle JDownloader2 Backup wieder her." -Level "INFO"
        Write-Host "	JDownloader2 Backup wird wiederhergestellt" -foregroundcolor "Yellow"
        try {
            & "C:\Program Files\7-Zip\7z.exe" x "$Serverip\Daten\Prog\Jdownloader_Backup.jd2backup" -o"C:\Program Files\JDownloader" -aoa
            Write_LogEntry -Message "Backup entpackt nach C:\Program Files\JDownloader von $($Serverip)\Daten\Prog\Jdownloader_Backup.jd2backup" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Entpacken des Backups: $($_)" -Level "ERROR"
        }
    } else {
        Write_LogEntry -Message "7-Zip nicht gefunden. Backup Wiederherstellung übersprungen." -Level "WARNING"
    }
}

$startMenuPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\JDownloader"
Write_LogEntry -Message "Prüfe Startmenü Pfad: $($startMenuPath)" -Level "DEBUG"
if (Test-Path $startMenuPath) {
    Write_LogEntry -Message "Startmenüeintrag gefunden: $($startMenuPath). Entferne Eintrag." -Level "INFO"
    Write-Host "	Startmenüeintrag wird entfernt." -foregroundcolor "Cyan"
    try {
        Remove-Item -Path $startMenuPath -Recurse -Force
        Write_LogEntry -Message "Startmenüeintrag entfernt: $($startMenuPath)" -Level "SUCCESS"
    } catch {
        Write_LogEntry -Message "Fehler beim Entfernen des Startmenüeintrags $($startMenuPath): $($_)" -Level "ERROR"
    }
} else {
    Write_LogEntry -Message "Kein Startmenüeintrag gefunden unter: $($startMenuPath)" -Level "DEBUG"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
