$ProgramName = "Security intelligence Update Kit"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

Write_LogEntry -Message "Script gestartet" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$NetworkShareDaten  = $config.NetworkShareDaten

# Constants
$webPageUrl    = "https://www.microsoft.com/en-us/wdsi/defenderupdates"
$downloadLink  = "https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64"
$directory     = "$NetworkShareDaten\Customize_Windows\Windows_Defender_Update_Iso\defender-update-kit-x64"
$cabFilePath   = "$directory\defender-dism-x64.cab"
$xmlFileName   = "package-defender.xml"
$extractPath   = "$env:TEMP\defender"
$sevenZipPath  = "C:\Program Files\7-Zip\7z.exe"
$makecabScript = "$NetworkShareDaten\Customize_Windows\Scripte\makecab.ps1"
$signerScript  = "$NetworkShareDaten\Customize_Windows\Scripte\certs\CreateCerts\Signer.ps1"

Write_LogEntry -Message "Konstanten: WebPageUrl=$($webPageUrl); Directory=$($directory); CabPath=$($cabFilePath); ExtractPath=$($extractPath)" -Level "DEBUG"

# --- Helper Functions ---

function New-DirectorySafe {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write_LogEntry -Message "Verzeichnis erstellt: $($Path)" -Level "INFO"
        } catch {
            Write_LogEntry -Message "Fehler beim Erstellen des Verzeichnisses $($Path): $($_)" -Level "ERROR"
            throw
        }
    } else {
        Write_LogEntry -Message "Verzeichnis bereits vorhanden: $($Path)" -Level "DEBUG"
    }
}

function Remove-DirectorySafe {
    param([string]$Path)
    if (Test-Path -Path $Path) {
        try {
            Remove-Item -Path $Path -Force -Recurse -ErrorAction Stop
            Write_LogEntry -Message "Verzeichnis gelöscht: $($Path)" -Level "INFO"
        } catch {
            Write-Host "Warnung: Konnte temporäre Dateien nicht löschen: $Path" -ForegroundColor "Yellow"
            Write_LogEntry -Message "Warnung: Fehler beim Löschen des Verzeichnisses $($Path): $($_)" -Level "WARNING"
        }
    } else {
        Write_LogEntry -Message "Zu löschendes Verzeichnis nicht vorhanden: $($Path)" -Level "DEBUG"
    }
}

function Invoke-CabExtraction {
    param(
        [string]$CabPath,
        [string]$DestPath
    )
    Write_LogEntry -Message "Starte CAB-Extraktion: $($CabPath) -> $($DestPath)" -Level "INFO"
    if (-not (Test-Path $sevenZipPath)) {
        Write_LogEntry -Message "7-Zip nicht gefunden: $($sevenZipPath)" -Level "ERROR"
        throw "7-Zip nicht gefunden: $sevenZipPath"
    }
    $result = & $sevenZipPath x -o"$DestPath" "$CabPath" -y 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write_LogEntry -Message "Fehler bei CAB-Extraktion (ExitCode $LASTEXITCODE): $($result)" -Level "ERROR"
        throw "Fehler bei CAB-Extraktion: $result"
    }
    Write_LogEntry -Message "CAB-Extraktion erfolgreich." -Level "SUCCESS"
}

function Invoke-7ZipExtract {
    param(
        [string]$SourceFile,
        [string]$DestPath
    )
    Write_LogEntry -Message "Starte 7-Zip Extraktion: $($SourceFile) -> $($DestPath)" -Level "INFO"
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo.FileName               = $sevenZipPath
    $proc.StartInfo.Arguments              = "x `"$SourceFile`" -o`"$DestPath`" -r -aoa"
    $proc.StartInfo.WindowStyle            = 'Hidden'
    $proc.StartInfo.UseShellExecute        = $false
    $proc.StartInfo.RedirectStandardOutput = $true
    $proc.Start() | Out-Null
    $proc.WaitForExit()
    $output = $proc.StandardOutput.ReadToEnd()
    $proc.Close()
    Write_LogEntry -Message "7-Zip Ausgabe: $($output.Trim())" -Level "DEBUG"
    if ($output -match "Everything is Ok") {
        Write_LogEntry -Message "7-Zip Extraktion erfolgreich: $($SourceFile)" -Level "SUCCESS"
        return $true
    } else {
        Write_LogEntry -Message "7-Zip Extraktion fehlgeschlagen oder unvollständig: $($SourceFile)" -Level "WARNING"
        return $false
    }
}

# --- Main Execution ---

try {
    Write-Host "Prüfe auf $ProgramName Updates..." -ForegroundColor "Cyan"
    Write_LogEntry -Message "Starte Hauptlogik" -Level "INFO"

    # Step 1: Fetch online versions
    try {
        Write_LogEntry -Message "Rufe Webseite ab: $($webPageUrl)" -Level "INFO"
        $webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing -TimeoutSec 30
        $content = $webPageContent.Content
        Write_LogEntry -Message "Webseite abgerufen; Content-Länge: $($content.Length)" -Level "DEBUG"
    } catch {
        Write-Host "Fehler beim Abrufen der Webseite: $_" -ForegroundColor "Red"
        Write_LogEntry -Message "Fehler beim Abrufen der Webseite $($webPageUrl): $($_)" -Level "ERROR"
        exit 1
    }

    $lineRegex = '(?s)<li>Version: <span>(\d+\.\d+\.\d+\.\d+)</span><\/li>\s+<li>Engine Version: <span>(\d+\.\d+\.\d+\.\d+)</span><\/li>\s+<li>Platform Version: <span>(\d+\.\d+\.\d+\.\d+)</span><\/li>'
    $lineMatch = [regex]::Match($content, $lineRegex)

    if (-not $lineMatch.Success) {
        Write-Host "Fehler: Online-Versionsinformationen konnten nicht extrahiert werden." -ForegroundColor "Red"
        Write_LogEntry -Message "Regex-Match für Online-Versionen fehlgeschlagen." -Level "ERROR"
        exit 1
    }

    $intelligenceVersionWeb = $lineMatch.Groups[1].Value
    $engineVersionWeb       = $lineMatch.Groups[2].Value
    $platformVersionWeb     = $lineMatch.Groups[3].Value
    Write_LogEntry -Message "Online-Versionen: Intelligence=$($intelligenceVersionWeb); Engine=$($engineVersionWeb); Platform=$($platformVersionWeb)" -Level "INFO"

    # Step 2: Extract local CAB versions
    New-DirectorySafe -Path $extractPath
    Invoke-CabExtraction -CabPath $cabFilePath -DestPath $extractPath

    $xmlFilePath = Join-Path $extractPath $xmlFileName
    if (-not (Test-Path $xmlFilePath)) {
        Write_LogEntry -Message "XML-Datei nicht in CAB gefunden: $($xmlFilePath)" -Level "ERROR"
        throw "XML-Datei '$xmlFileName' nicht in der CAB-Datei gefunden."
    }

    $xml = [xml](Get-Content -Path $xmlFilePath -Raw)
    $engineVersionCab     = $xml.packageinfo.versions.engine
    $platformVersionCab   = $xml.packageinfo.versions.platform
    $signaturesVersionCab = $xml.packageinfo.versions.signatures
    Write_LogEntry -Message "Lokale Versionen: Intelligence=$($signaturesVersionCab); Engine=$($engineVersionCab); Platform=$($platformVersionCab)" -Level "INFO"

    # Step 3: Display and compare
    $cabFileVersions = @{
        "Platform version"              = $platformVersionCab
        "Engine version"                = $engineVersionCab
        "Security intelligence version" = $signaturesVersionCab
    }
    $webPageVersions = @{
        "Platform version"              = $platformVersionWeb
        "Engine version"                = $engineVersionWeb
        "Security intelligence version" = $intelligenceVersionWeb
    }

    $cabFileVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Local";  Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize
    $webPageVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Online"; Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize

    $UpdateAvailable = $false
    Write-Host "Vergleich:" -ForegroundColor "DarkGray"
    foreach ($key in $webPageVersions.Keys) {
        $webVersion = $webPageVersions[$key]
        $cabVersion = $cabFileVersions[$key]
        try {
            if ([version]$webVersion -gt [version]$cabVersion) {
                Write-Host "   $key ist unterschiedlich:" -ForegroundColor "DarkGray"
                Write-Host "      Lokal:  $cabVersion" -ForegroundColor "Cyan"
                Write-Host "      Online: $webVersion" -ForegroundColor "Cyan"
                $UpdateAvailable = $true
                Write_LogEntry -Message "Update für $($key): Online $($webVersion) > Lokal $($cabVersion)" -Level "INFO"
            } else {
                Write-Host "   $key stimmt überein. Version: $webVersion" -ForegroundColor "DarkGray"
                Write_LogEntry -Message "$($key) aktuell: $($webVersion)" -Level "DEBUG"
            }
        } catch {
            Write-Host "   $key - Fehler beim Versionsvergleich" -ForegroundColor "Yellow"
            Write_LogEntry -Message "Fehler beim Vergleich von $($key): $($_)" -Level "ERROR"
        }
    }

    Write-Host ""

    if (-not $UpdateAvailable) {
        Write-Host "Kein Update verfügbar. $ProgramName ist aktuell." -ForegroundColor "DarkGray"
        Write_LogEntry -Message "Kein Update verfügbar für $($ProgramName)" -Level "INFO"
    } else {
        Write-Host "Update ist verfügbar." -ForegroundColor "Green"
        Write_LogEntry -Message "Update verfügbar. Starte Update-Prozess." -Level "INFO"
        Write-Host ""

        # Step 4: Download signatures
        $tempFilePath = Join-Path $env:TEMP "mpam-fe.exe"
        Write_LogEntry -Message "Lade Signaturen herunter: $($downloadLink) -> $($tempFilePath)" -Level "INFO"
        Write-Host "Lade Signaturen herunter..." -ForegroundColor "Yellow"
        try {
            $webClient = New-Object System.Net.WebClient
            [void](Invoke-DownloadFile -Url $downloadLink -OutFile $tempFilePath)
            $webClient.Dispose()
            Write-Host "Download abgeschlossen." -ForegroundColor "Green"
            Write_LogEntry -Message "Download abgeschlossen: $($tempFilePath)" -Level "SUCCESS"
        } catch {
            Write_LogEntry -Message "Fehler beim Herunterladen: $($_)" -Level "ERROR"
            throw "Download fehlgeschlagen: $_"
        }

        # Step 5: Extract signatures
        $extractPathSub = Join-Path $extractPath "Definition Updates\Updates"
        New-DirectorySafe -Path $extractPathSub

        $extractOk = Invoke-7ZipExtract -SourceFile $tempFilePath -DestPath $extractPathSub
        if (-not $extractOk) {
            Write_LogEntry -Message "Extraktion von mpam-fe.exe fehlgeschlagen." -Level "ERROR"
            throw "Extraktion von mpam-fe.exe fehlgeschlagen."
        }

        # Remove MpSigStub.exe
        $mpSigStubPath = Join-Path $extractPathSub "MpSigStub.exe"
        if (Test-Path $mpSigStubPath) {
            Remove-Item -Path $mpSigStubPath -Force
            Write_LogEntry -Message "MpSigStub.exe entfernt." -Level "DEBUG"
        }

        # Step 6: Update platform if needed
        $updatePlatformFileVersion = $null
        if ($platformVersionCab -ne $platformVersionWeb) {
            Write_LogEntry -Message "Plattform-Update erforderlich: Lokal=$($platformVersionCab); Online=$($platformVersionWeb)" -Level "INFO"

            $moduleName = "kbupdate"
            if (-not (Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName })) {
                Write_LogEntry -Message "Installiere Modul: $($moduleName)" -Level "INFO"
                Install-Module -Name $moduleName -Force -Scope CurrentUser -AllowClobber
            }
            Import-Module $moduleName -ErrorAction Stop
            Write_LogEntry -Message "Modul '$($moduleName)' geladen." -Level "DEBUG"

            $updateList   = Get-KbUpdate -Name KB4052623
            $newestUpdate = $updateList | Sort-Object -Property LastModified -Descending | Select-Object -First 1

            if (-not $newestUpdate) {
                Write-Host "Update KB4052623 nicht im Windows Update Katalog gefunden." -ForegroundColor "Red"
                Write_LogEntry -Message "KB4052623 nicht gefunden." -Level "ERROR"
                throw "KB4052623 nicht gefunden."
            }

            $updateLink = $newestUpdate.Link | Where-Object { $_ -match "updateplatform\.amd64fre" }
            if (-not $updateLink) {
                Write-Host "Update-Link 'updateplatform.amd64fre' nicht gefunden." -ForegroundColor "Red"
                Write_LogEntry -Message "Update-Link nicht gefunden in KB4052623." -Level "ERROR"
                throw "Plattform-Update-Link nicht gefunden."
            }

            $downloadPath = Join-Path $env:TEMP ($updateLink.Split("/")[-1])
            Write_LogEntry -Message "Lade Plattform-Update herunter: $($updateLink) -> $($downloadPath)" -Level "INFO"
            [void](Invoke-DownloadFile -Url $updateLink -OutFile $downloadPath)

            # The downloaded file itself may be the UpdatePlatform exe
            $updatePlatformFile = $null
            if ((Split-Path $downloadPath -Leaf) -match "UpdatePlatform") {
                $updatePlatformFile = Get-Item $downloadPath -ErrorAction SilentlyContinue
            }
            if (-not $updatePlatformFile) {
                $updatePlatformFile = Get-ChildItem -Path (Split-Path $downloadPath) -Filter "UpdatePlatform*.exe" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            }

            if ($updatePlatformFile) {
                $updatePlatformFileVersion = (Get-Item $updatePlatformFile.FullName).VersionInfo.ProductVersion
                Write_LogEntry -Message "Plattform-Datei: $($updatePlatformFile.Name); Version: $($updatePlatformFileVersion)" -Level "DEBUG"

                if ($platformVersionCab -ne $updatePlatformFileVersion) {
                    $platformPath  = Join-Path $extractPath "Platform"
                    $currentFolder = (Get-ChildItem -Path $platformPath -Directory -ErrorAction SilentlyContinue | Select-Object -Last 1).Name
                    if ($currentFolder) {
                        $oldFolderPath = Join-Path $platformPath $currentFolder
                        Remove-Item -Path $oldFolderPath -Force -Recurse
                        Write_LogEntry -Message "Alten Plattform-Ordner entfernt: $($oldFolderPath)" -Level "DEBUG"
                    }

                    $newFolderPath = Join-Path $platformPath "$platformVersionWeb-0"
                    Write_LogEntry -Message "Extrahiere Plattform nach: $($newFolderPath)" -Level "INFO"

                    $proc = New-Object System.Diagnostics.Process
                    $proc.StartInfo.FileName               = $sevenZipPath
                    $proc.StartInfo.Arguments              = "x `"$($updatePlatformFile.FullName)`" -o`"$newFolderPath`" -r"
                    $proc.StartInfo.WindowStyle            = 'Hidden'
                    $proc.StartInfo.UseShellExecute        = $false
                    $proc.StartInfo.RedirectStandardOutput = $true
                    $proc.Start() | Out-Null
                    $proc.WaitForExit()
                    $proc.Close()
                    Write_LogEntry -Message "Plattform-Extraktion abgeschlossen: $($newFolderPath)" -Level "SUCCESS"

                    Remove-Item -Path $updatePlatformFile.FullName -Force
                    Write_LogEntry -Message "Plattform-Update-Datei entfernt: $($updatePlatformFile.FullName)" -Level "DEBUG"
                }
            } else {
                Write_LogEntry -Message "UpdatePlatform.exe nicht gefunden nach Download." -Level "WARNING"
            }
        }

        # Step 7: Update package-defender.xml
        $packageXmlFile = Join-Path $extractPath "package-defender.xml"
        Write_LogEntry -Message "Aktualisiere XML: $($packageXmlFile)" -Level "INFO"
        $xml = [xml](Get-Content $packageXmlFile)
        $xml.packageinfo.versions.signatures = $intelligenceVersionWeb
        $xml.packageinfo.versions.engine     = $engineVersionWeb
        if ($updatePlatformFileVersion -and ($platformVersionCab -ne $updatePlatformFileVersion)) {
            $xml.packageinfo.versions.platform = $platformVersionWeb
            Write_LogEntry -Message "Plattform-Version in XML aktualisiert: $($platformVersionWeb)" -Level "INFO"
        }
        $xml.Save($packageXmlFile)
        Write_LogEntry -Message "XML gespeichert: $($packageXmlFile)" -Level "SUCCESS"

        # Cleanup mpam-fe.exe
        Remove-Item -Path $tempFilePath -Force -ErrorAction SilentlyContinue
        Write_LogEntry -Message "mpam-fe.exe entfernt." -Level "DEBUG"

        # Step 8: Create new CAB
        Write_LogEntry -Message "Starte CAB-Erstellung aus: $($extractPath)" -Level "INFO"
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
        . $makecabScript
        New-CAB -SourceDirectory $extractPath

        $cabFile = Get-ChildItem -Path $env:TEMP -Filter "*.cab" -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $cabFile) {
            Write_LogEntry -Message "Keine CAB-Datei im TEMP nach Erstellung gefunden." -Level "ERROR"
            throw "CAB-Erstellung fehlgeschlagen."
        }

        $finalCabPath = Join-Path $env:TEMP "defender-dism-x64.cab"
        if ($cabFile.FullName -ne $finalCabPath) {
            Move-Item -Path $cabFile.FullName -Destination $finalCabPath -Force
            Write_LogEntry -Message "CAB umbenannt nach: $($finalCabPath)" -Level "DEBUG"
        }

        # Delete DDF temp files
        Get-ChildItem -Path $env:TEMP -Filter "*.ddf" -File | Remove-Item -Force
        Write_LogEntry -Message "DDF-Dateien bereinigt." -Level "DEBUG"

        # Step 9: Sign the CAB
        if (-not (Test-Path $signerScript)) {
            Write_LogEntry -Message "Signer-Skript nicht gefunden: $($signerScript)" -Level "ERROR"
            throw "Signer-Skript nicht gefunden: $signerScript"
        }
        try { Unblock-File -Path $signerScript -ErrorAction SilentlyContinue } catch {}

        Write-Host "Starte Signer..." -ForegroundColor "Yellow"
        Write_LogEntry -Message "Starte Signer: $($signerScript) mit Argument: $($finalCabPath)" -Level "INFO"

        $winPs    = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
        $output   = & $winPs -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $signerScript $finalCabPath 2>&1
        $exitCode = $LASTEXITCODE

        if ($output) {
            $outText = if ($output -is [array]) { $output -join "`n" } else { [string]$output }
            Write_LogEntry -Message "Signer-Ausgabe: $($outText)" -Level "DEBUG"
        }

        if ($exitCode -ne 0) {
            Write_LogEntry -Message "Signer fehlgeschlagen (ExitCode $($exitCode))" -Level "ERROR"
            throw "Signer fehlgeschlagen mit ExitCode $exitCode"
        }
        Write_LogEntry -Message "Signer erfolgreich (ExitCode 0)" -Level "SUCCESS"

        # Step 10: Move final CAB to destination
        Move-Item -Path $finalCabPath -Destination $directory -Force
        Write_LogEntry -Message "CAB nach $($directory) verschoben." -Level "SUCCESS"

        Write-Host ""
        Write-Host "$ProgramName wurde aktualisiert." -ForegroundColor "Green"
        Write_LogEntry -Message "$($ProgramName) Update erfolgreich abgeschlossen." -Level "SUCCESS"
    }
}
catch {
    Write-Host "Fehler: $_" -ForegroundColor "Red"
    Write_LogEntry -Message "Unbehandelter Fehler: $($_)" -Level "ERROR"
    exit 1
}
finally {
    Remove-DirectorySafe -Path $extractPath
    Write_LogEntry -Message "Aufräumen abgeschlossen: $($extractPath)" -Level "DEBUG"
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
