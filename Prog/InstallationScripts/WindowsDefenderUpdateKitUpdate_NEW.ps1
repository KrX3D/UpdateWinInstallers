$ProgramName = "Windows Defender Update Kit"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath

# Constants
$webPageUrl = "https://support.microsoft.com/en-us/topic/microsoft-defender-update-for-windows-operating-system-installation-images-1c89630b-61ff-00a1-04e2-2d1f3865450d"
$downloadLink = "https://go.microsoft.com/fwlink/?linkid=2144531"
$directory = "$NetworkShareDaten\Customize_Windows\Windows_Defender_Update_Iso\defender-update-kit-x64"
$cabFilePath = "$directory\defender-dism-x64.cab"
$xmlFileName = "package-defender.xml"
$extractPath = "$env:TEMP\defender"
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

Write_LogEntry -Message "Konstanten: WebPageUrl=$($webPageUrl); DownloadLink=$($downloadLink); Directory=$($directory); CabPath=$($cabFilePath); XmlFile=$($xmlFileName); ExtractPath=$($extractPath); 7Zip=$($sevenZipPath)" -Level "DEBUG"

# Regex patterns for version extraction
$versionPatterns = @{
    "Defender package version" = 'Defender package version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
    "Platform version" = 'Platform version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
    "Engine version" = 'Engine version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
    "Security intelligence version" = 'Security intelligence version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
}

# Function to extract version from the given HTML line using the specified pattern
function Get-VersionFromLine {
    param(
        [string]$Line,
        [string]$Pattern
    )
    
    if ([string]::IsNullOrEmpty($Line)) {
        return "Not found"
    }
    
    $match = [regex]::Match($Line, $Pattern)
    if ($match.Success) {
        return $match.Groups[1].Value
    } else {
        return "Not found"
    }
}

# Function to safely create directory
function New-DirectorySafe {
    param([string]$Path)
    
    if (-not (Test-Path -Path $Path)) {
        try {
            Write_LogEntry -Message "Erstelle Verzeichnis: $($Path)" -Level "DEBUG"
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write_LogEntry -Message "Verzeichnis erstellt: $($Path)" -Level "INFO"
        }
        catch {
            Write-Host "Fehler beim Erstellen des Verzeichnisses: $Path" -ForegroundColor "Red"
            Write_LogEntry -Message "Fehler beim Erstellen des Verzeichnisses $($Path): $($_)" -Level "ERROR"
            throw
        }
    } else {
        Write_LogEntry -Message "Verzeichnis bereits vorhanden: $($Path)" -Level "DEBUG"
    }
}

# Function to safely remove directory
function Remove-DirectorySafe {
    param([string]$Path)
    
    if (Test-Path -Path $Path) {
        try {
            Write_LogEntry -Message "Lösche Verzeichnis: $($Path)" -Level "DEBUG"
            Remove-Item -Path $Path -Force -Recurse -ErrorAction Stop
            Write_LogEntry -Message "Verzeichnis gelöscht: $($Path)" -Level "INFO"
        }
        catch {
            Write-Host "Warnung: Konnte temporäre Dateien nicht löschen: $Path" -ForegroundColor "Yellow"
            Write_LogEntry -Message "Warnung: Fehler beim Löschen des Verzeichnisses $($Path): $($_)" -Level "WARNING"
        }
    } else {
        Write_LogEntry -Message "Zu löschendes Verzeichnis nicht vorhanden: $($Path)" -Level "DEBUG"
    }
}

# Function to extract CAB file
function Invoke-CabExtraction {
    param(
        [string]$CabPath,
        [string]$ExtractPath
    )
    
    Write_LogEntry -Message "Starte CAB-Extraktion: CabPath=$($CabPath); ExtractPath=$($ExtractPath)" -Level "INFO"
    if (-not (Test-Path $sevenZipPath)) {
        Write_LogEntry -Message "7-Zip nicht gefunden unter: $($sevenZipPath)" -Level "ERROR"
        throw "7-Zip nicht gefunden unter: $sevenZipPath"
    }
    
    $result = & $sevenZipPath x -o"$ExtractPath" "$CabPath" -y 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write_LogEntry -Message "Fehler beim Extrahieren der CAB-Datei: $($result)" -Level "ERROR"
        throw "Fehler beim Extrahieren der CAB-Datei: $result"
    } else {
        Write_LogEntry -Message "CAB-Extraktion erfolgreich: $($CabPath) -> $($ExtractPath)" -Level "SUCCESS"
    }
}

# Function to download file with better error handling
function Invoke-FileDownload {
    param(
        [string]$Url,
        [string]$FilePath
    )
    
    try {
        Write-Host "Lade Datei herunter..." -ForegroundColor "Yellow"
        Write_LogEntry -Message "Starte Download: $($Url) -> $($FilePath)" -Level "INFO"
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $FilePath)
        $webClient.Dispose()
        Write-Host "Download abgeschlossen." -ForegroundColor "Green"
        Write_LogEntry -Message "Download abgeschlossen: $($FilePath)" -Level "SUCCESS"
    }
    catch {
        Write_LogEntry -Message "Fehler beim Herunterladen $($Url) nach $($FilePath): $($_)" -Level "ERROR"
        throw "Fehler beim Herunterladen: $_"
    }
}

# Function to modify PS1 file
function Edit-DefenderScript {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        Write-Host "Warnung: DefenderUpdateWinImage.ps1 nicht gefunden" -ForegroundColor "Yellow"
        Write_LogEntry -Message "DefenderUpdateWinImage.ps1 nicht gefunden: $($FilePath)" -Level "WARNING"
        return
    }
    
    try {
        Write_LogEntry -Message "Versuche DefenderUpdateWinImage.ps1 zu bearbeiten: $($FilePath)" -Level "DEBUG"
        $fileContent = Get-Content -Path $FilePath
        $lineToComment = '    ValidateCodeSign -PackageFile $PkgFile'
        
        $modified = $false
        for ($i = 0; $i -lt $fileContent.Length; $i++) {
            if ($fileContent[$i] -eq $lineToComment) {
                $fileContent[$i] = '#' + $lineToComment
                $modified = $true
                break
            }
        }
        
        if ($modified) {
            $fileContent | Set-Content -Path $FilePath
            Write_LogEntry -Message "DefenderUpdateWinImage.ps1 wurde modifiziert: $($FilePath)" -Level "INFO"
        } else {
            Write_LogEntry -Message "Keine Änderung erforderlich in DefenderUpdateWinImage.ps1: $($FilePath)" -Level "DEBUG"
        }
    }
    catch {
        Write-Host "Warnung: Konnte DefenderUpdateWinImage.ps1 nicht bearbeiten: $_" -ForegroundColor "Yellow"
        Write_LogEntry -Message "Warnung beim Bearbeiten von DefenderUpdateWinImage.ps1 $($FilePath): $($_)" -Level "WARNING"
    }
}

# Main execution
try {
    Write-Host "Prüfe auf $ProgramName Updates..." -ForegroundColor "Cyan"
    Write_LogEntry -Message "Prüfe auf $($ProgramName) Updates - Starte Hauptlogik" -Level "INFO"
    
    # Get web page content
    try {
        Write_LogEntry -Message "Hole Webseite: $($webPageUrl)" -Level "INFO"
        $webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing -TimeoutSec 30
        Write_LogEntry -Message "Webseite abgerufen; Content-Länge: $($webPageContent.Content.Length)" -Level "DEBUG"
    }
    catch {
        Write-Host "Fehler beim Abrufen der Webseite: $_" -ForegroundColor "Red"
        Write_LogEntry -Message "Fehler beim Abrufen der Webseite $($webPageUrl): $($_)" -Level "ERROR"
        exit 1
    }
    
    # Extract version information from web page
    $webPageVersions = @{}
    foreach ($key in $versionPatterns.Keys) {
        $line = $webPageContent.Content | Select-String -Pattern $key
        $webPageVersions[$key] = Get-VersionFromLine $line $versionPatterns[$key]
        Write_LogEntry -Message "Ermittelte Online-Version für $($key): $($webPageVersions[$key])" -Level "DEBUG"
    }
    
    # Check if CAB file exists
    if (-not (Test-Path $cabFilePath)) {
        Write-Host "Lokale CAB-Datei nicht gefunden. Führe ersten Download durch..." -ForegroundColor "Yellow"
        Write_LogEntry -Message "Lokale CAB-Datei nicht gefunden: $($cabFilePath)" -Level "WARNING"
        $UpdateAvailable = $true
        $cabFileVersions = @{
            "Defender package version" = "0.0.0.0"
            "Platform version" = "0.0.0.0"
            "Engine version" = "0.0.0.0"
            "Security intelligence version" = "0.0.0.0"
        }
    } else {
        Write_LogEntry -Message "Lokale CAB-Datei gefunden: $($cabFilePath) - Starte Extraktion" -Level "INFO"
        # Create extraction directory and extract CAB
        New-DirectorySafe -Path $extractPath
        Invoke-CabExtraction -CabPath $cabFilePath -ExtractPath $extractPath
        
        # Read XML file
        $xmlFilePath = Join-Path $extractPath $xmlFileName
        if (-not (Test-Path $xmlFilePath)) {
            Write_LogEntry -Message "XML-Datei '$($xmlFileName)' nicht in CAB gefunden: $($xmlFilePath)" -Level "ERROR"
            throw "XML-Datei '$xmlFileName' nicht in der CAB-Archive gefunden."
        }
        
        $xmlContent = Get-Content -Path $xmlFilePath -Raw
        $xml = [xml]$xmlContent
        
        # Extract versions from XML
        $cabFileVersions = @{
            "Defender package version" = $xml.packageinfo.versions.defender
            "Platform version" = $xml.packageinfo.versions.platform
            "Engine version" = $xml.packageinfo.versions.engine
            "Security intelligence version" = $xml.packageinfo.versions.signatures
        }
		
		$cabPairsString = ($cabFileVersions.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
		Write_LogEntry -Message "Versionen aus CAB ermittelt: $($cabPairsString)" -Level "DEBUG"

        # Compare versions (but don't display yet)
        $UpdateAvailable = $false
        $comparisonResults = @()
        
        foreach ($key in $webPageVersions.Keys) {
            $webVersion = $webPageVersions[$key]
            $cabVersion = $cabFileVersions[$key]
            
            if ($webVersion -eq "Not found" -or $cabVersion -eq $null) {
                $comparisonResults += "   ${key} - Versionsinformation unvollständig"
                Write_LogEntry -Message "Versionsinformation unvollständig für $($key): Web=$($webVersion); Cab=$($cabVersion)" -Level "WARNING"
                continue
            }
            
            try {
                if ([version]$webVersion -gt [version]$cabVersion) {
                    $comparisonResults += "   ${key} ist unterschiedlich:"
                    $comparisonResults += "      Lokal: $cabVersion"
                    $comparisonResults += "      Online: $webVersion"
                    $UpdateAvailable = $true
                    Write_LogEntry -Message "Update für $($key) verfügbar: Online $($webVersion) > Local $($cabVersion)" -Level "INFO"
                } else {
                    $comparisonResults += "   ${key} stimmt überein. Version: $webVersion"
                    Write_LogEntry -Message "$($key) stimmt überein: $($webVersion)" -Level "DEBUG"
                }
            }
            catch {
                $comparisonResults += "   ${key} - Fehler beim Vergleichen der Versionen"
                Write_LogEntry -Message "Fehler beim Vergleichen der Versionen für $($key): $($_)" -Level "ERROR"
            }
        }
    }
    
    # Display versions in correct order: Local, Online, then Comparison
    if (Test-Path $cabFilePath) {
        $cabFileVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Local"; Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize
    }
    
    $webPageVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Online"; Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize
    
    # Display comparison results
    if (Test-Path $cabFilePath) {
        Write-Host "Vergleich:" -ForegroundColor "DarkGray"
        foreach ($result in $comparisonResults) {
            if ($result -like "*unterschiedlich*") {
                Write-Host $result -ForegroundColor "DarkGray"
            } elseif ($result -like "*Lokal:*") {
                Write-Host $result -ForegroundColor "Cyan"
            } elseif ($result -like "*Online:*") {
                Write-Host $result -ForegroundColor "Cyan"
            } elseif ($result -like "*stimmt überein*") {
                Write-Host $result -ForegroundColor "DarkGray"
            } else {
                Write-Host $result -ForegroundColor "Yellow"
            }
        }
    }
    
    # Process update if available
    if ($UpdateAvailable) {
        Write-Host "Update ist verfügbar." -ForegroundColor "Green"
        Write_LogEntry -Message "Update verfügbar: Starte Update-Prozess" -Level "INFO"
        Write-Host ""
        
        $tempFilePath = Join-Path $env:TEMP "defender-update-kit-x64.cab"
        Write_LogEntry -Message "Temporäre CAB-Datei Pfad: $($tempFilePath)" -Level "DEBUG"
        
        # Download new version
        Invoke-FileDownload -Url $downloadLink -FilePath $tempFilePath
        
        # Create target directory and clean old files
        New-DirectorySafe -Path $directory
        if (Test-Path "$directory\*") {
            Write_LogEntry -Message "Bereinige Zielverzeichnis: $($directory)" -Level "DEBUG"
            Remove-Item -Path "$directory\*" -Force -Recurse
        }
        
        # Extract downloaded file
        Invoke-CabExtraction -CabPath $tempFilePath -ExtractPath $directory
        
        # Clean up downloaded file
        Remove-Item -Path $tempFilePath -Force -ErrorAction SilentlyContinue
        Write_LogEntry -Message "Temporäre Datei gelöscht: $($tempFilePath)" -Level "DEBUG"
        
        # Edit downloaded PS1 file to skip signature check
        $defenderScriptPath = "$directory\DefenderUpdateWinImage.ps1"
        Edit-DefenderScript -FilePath $defenderScriptPath
        
        Write-Host ""
        Write-Host "$ProgramName wurde aktualisiert." -ForegroundColor "Green"
        Write_LogEntry -Message "$($ProgramName) Update erfolgreich durchgeführt in $($directory)" -Level "SUCCESS"
    } else {
        Write-Host ""
        Write-Host "Kein Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
        Write_LogEntry -Message "Kein Update verfügbar für $($ProgramName)" -Level "INFO"
    }
}
catch {
    Write-Host "Fehler: $_" -ForegroundColor "Red"
    Write_LogEntry -Message "Unhandled error: $($_)" -Level "ERROR"
    exit 1
}
finally {
    # Clean up extracted files
    Remove-DirectorySafe -Path $extractPath
    Write_LogEntry -Message "Aufräumen abgeschlossen: $($extractPath)" -Level "DEBUG"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
