param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinSCP"
$ScriptType  = "Update"

# Determine whether to add -UseBasicParsing (only valid in Windows PowerShell <= 5.1)
$psMajor = if ($PSVersionTable -and $PSVersionTable.PSVersion) { $PSVersionTable.PSVersion.Major } else { 5 }
$UseBasicParsingSupported = $false
if ($psMajor -lt 6) { $UseBasicParsingSupported = $true }

function Invoke-WebRequestCompat {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [string]$OutFile = $null
    )
    if ($UseBasicParsingSupported) {
        if ($OutFile) {
            return (Invoke-DownloadFile -Url $Uri -OutFile $OutFile)
        } else {
            return Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop
        }
    } else {
        if ($OutFile) {
            return (Invoke-DownloadFile -Url $Uri -OutFile $OutFile)
        } else {
            return Invoke-WebRequest -Uri $Uri -ErrorAction Stop
        }
    }
}

# === Logger-Header: automatisch eingefügt ===
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\Logger\Logger.psm1"

if (Test-Path $modulePath) {
    Import-Module -Name $modulePath -Force -ErrorAction Stop

    if (-not (Get-Variable -Name logRoot -Scope Script -ErrorAction SilentlyContinue)) {
        $logRoot = Join-Path -Path $PSScriptRoot -ChildPath "Log"
    }
    Set_LoggerConfig -LogRootPath $logRoot | Out-Null

    if (Get-Command -Name Initialize_LogSession -ErrorAction SilentlyContinue) {
        Initialize_LogSession -ProgramName $ProgramName -ScriptType $ScriptType | Out-Null #-WriteSystemInfo
    }
}
# === Ende Logger-Header ===

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType); PS Major: $psMajor; UseBasicParsingSupported: $UseBasicParsingSupported" -Level "DEBUG"

# DeployToolkit helpers
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
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
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Write_LogEntry -Message "Berechneter Konfigurationspfad: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei geladen: $($configPath)" -Level "INFO"
} else {
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    exit
}

$installerPath = "$InstallationFolder\WinSCP-*.exe"
Write_LogEntry -Message "Installer-Pfad (Wildcard): $($installerPath)" -Level "DEBUG"

$InstallationFileFile = Get-ChildItem -Path $installerPath -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($InstallationFileFile) {
    Write_LogEntry -Message "Gefundene Installationsdatei: $($InstallationFileFile.FullName)" -Level "INFO"
} else {
    Write_LogEntry -Message "Keine Installationsdatei gefunden im Pfad: $($installerPath)" -Level "WARNING"
}

# Check if the installer file exists
if ($InstallationFileFile) {
    # Extract the version number from the installer properties
    $versionInfo = (Get-Item $InstallationFileFile).VersionInfo
    $localVersion = ($versionInfo.ProductVersion).ToString().Trim()
    Write_LogEntry -Message "Lokale Installationsdatei Version ermittelt: $($localVersion) aus Datei $($InstallationFileFile.FullName)" -Level "DEBUG"

    # Retrieve the latest version online from the WinSCP website
    $webPageUrl = "https://winscp.net/eng/downloads.php"
    Write_LogEntry -Message "Hole Web-Seite zur Versionsermittlung: $($webPageUrl)" -Level "INFO"
    try {
        $webPageContent = Invoke-WebRequestCompat -Uri $webPageUrl
        if ($null -ne $webPageContent -and $null -ne $webPageContent.Content) {
            Write_LogEntry -Message "Webseite abgerufen: $($webPageUrl); Content-Länge: $($webPageContent.Content.Length)" -Level "DEBUG"
        } else {
            Write_LogEntry -Message "Webseite abgerufen, aber Content leer oder nicht verfügbar." -Level "DEBUG"
        }
    } catch {
        Write_LogEntry -Message "Fehler beim Abrufen der Webseite $($webPageUrl): $($_)" -Level "ERROR"
        $webPageContent = $null
    }

    # parse latest version from page (robust-ish)
    $latestVersion = $null

    # Helper: try to extract version from any content blob (main page, redirect page, link hrefs)
    function Get-VersionFromFilenameInContent {
        param($content)
        if (-not $content) { return $null }
        $m = [regex]::Match($content, 'WinSCP-([0-9]+(?:\.[0-9]+)+)-Setup\.exe', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success -and $m.Groups.Count -ge 2) { return $m.Groups[1].Value.Trim() }
        return $null
    }

    # First try main page content for a filename
    if ($webPageContent -and $webPageContent.Content) {
        $latestVersion = Get-VersionFromFilenameInContent -content $webPageContent.Content
        if ($latestVersion) {
            Write_LogEntry -Message "Version aus Hauptseite (Dateiname) ermittelt: $($latestVersion)" -Level "DEBUG"
        }
    }

    # If not found, try to find candidate links on the main page (Links property)
    if (-not $latestVersion -and $webPageContent -and $webPageContent.Links) {
        try {
            foreach ($lnk in $webPageContent.Links) {
                if ($lnk.href) {
                    $v = Get-VersionFromFilenameInContent -content $lnk.href
                    if ($v) { $latestVersion = $v; break }
                }
                if ($lnk.innerText) {
                    $v = Get-VersionFromFilenameInContent -content $lnk.innerText
                    if ($v) { $latestVersion = $v; break }
                }
            }
            if ($latestVersion) { Write_LogEntry -Message "Version aus Hauptseiten-Links ermittelt: $($latestVersion)" -Level "DEBUG" }
        } catch {
            Write_LogEntry -Message "Fehler beim Durchsuchen von Links auf der Hauptseite: $($_)" -Level "DEBUG"
        }
    }

    # If still not found, fall back to older patterns on the main page
    if (-not $latestVersion -and $webPageContent -and $webPageContent.Content) {
        $patterns = @(
            'Download\s*<strong>\s*WinSCP\s*<\/strong>\s*([0-9]+(?:\.[0-9]+)+)', # Download <strong>WinSCP</strong> 6.5.4
            'WinSCP\s*([0-9]+(?:\.[0-9]+)+)',                                     # WinSCP 6.5.4 (looser)
            'Version:\s*([0-9]+(?:\.[0-9]+)+)'                                   # Version: x.y.z
        )
        foreach ($p in $patterns) {
            $m = [regex]::Match($webPageContent.Content, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($m.Success -and $m.Groups.Count -ge 2) {
                $latestVersion = $m.Groups[1].Value.Trim()
                Write_LogEntry -Message "Version (Fallback pattern) ermittelt: $($latestVersion) via pattern $p" -Level "DEBUG"
                break
            }
        }
    }

    # If we have something that looks incomplete (e.g. only two components like "6.5"), try the redirect page to get filename-based version
    if ($latestVersion -and ($latestVersion.Split('.').Length -lt ($localVersion.Split('.').Length))) {
        Write_LogEntry -Message "Gefundene Online-Version '$latestVersion' sieht unvollständig aus (weniger Komponenten als lokale Version). Versuche Redirect-Seite/links zur exakten Version." -Level "DEBUG"
        $latestVersion = $null
    }

    # If still not found, probe the redirect URL for the expected setup filename (this often contains the full version)
    if (-not $latestVersion) {
        $testRedirectUrl = "https://winscp.net/download/WinSCP-*-Setup.exe"
        # We'll attempt the specific redirect for the latest guess (if we have one) or try to grab from the known download landing
        try {
            # Try the generic download page which often has the final link
            $downloadLanding = Invoke-WebRequestCompat -Uri "https://winscp.net/download/" -ErrorAction Stop
            if ($downloadLanding -and $downloadLanding.Content) {
                $latestVersion = Get-VersionFromFilenameInContent -content $downloadLanding.Content
                if ($latestVersion) {
                    Write_LogEntry -Message "Version aus download-Landing ermittelt: $($latestVersion)" -Level "DEBUG"
                }
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Abrufen der Download-Landing-Seite: $($_)" -Level "DEBUG"
        }
    }

    if (-not $latestVersion) {
        Write_LogEntry -Message "Konnte Online-Version nicht zuverlässig extrahieren aus $($webPageUrl)." -Level "WARNING"
        $latestVersion = ""
    } else {
        Write_LogEntry -Message "Ermittelte Online-Version: $($latestVersion)" -Level "DEBUG"
    }

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -foregroundcolor "Cyan"
    Write-Host "Online Version: $latestVersion" -foregroundcolor "Cyan"
    Write-Host ""
    Write_LogEntry -Message "Vergleich Local: $($localVersion) vs Online: $($latestVersion)" -Level "INFO"

    # safe version comparison helper
    function Test-IsVersionGreater {
        param([string]$a, [string]$b)
        if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { return $false }
        try {
            $va = [version]$a
            $vb = [version]$b
            return $va -gt $vb
        } catch {
            # fallback to numeric-ish compare: split by . and compare elements
            $as = $a.Split('.') | ForEach-Object { try { [int]$_ } catch { 0 } }
            $bs = $b.Split('.') | ForEach-Object { try { [int]$_ } catch { 0 } }
            $len = [Math]::Max($as.Length, $bs.Length)
            for ($i = 0; $i -lt $len; $i++) {
                $ai = if ($i -lt $as.Length) { $as[$i] } else { 0 }
                $bi = if ($i -lt $bs.Length) { $bs[$i] } else { 0 }
                if ($ai -gt $bi) { return $true }
                if ($ai -lt $bi) { return $false }
            }
            return $false
        }
    }

    # Compare the installed version with the latest version
    if ($latestVersion -and (Test-IsVersionGreater -a $latestVersion $localVersion)) {
        Write_LogEntry -Message "Online-Version größer als lokale Version: Update verfügbar (Online $($latestVersion) > Local $($localVersion))" -Level "INFO"

        # build the redirect/download starter URL (this is how you previously did it)
        $redirectUrl = "https://winscp.net/download/WinSCP-$latestVersion-Setup.exe"
        Write_LogEntry -Message "Hole Redirect-Seite: $($redirectUrl) um Download-Link zu ermitteln" -Level "DEBUG"

        try {
            $redirectPage = Invoke-WebRequestCompat -Uri $redirectUrl
            if ($redirectPage -and $redirectPage.Content) {
                Write_LogEntry -Message "Redirect-Seite abgerufen; Länge: $($redirectPage.Content.Length)" -Level "DEBUG"
            } else {
                Write_LogEntry -Message "Redirect-Seite abgerufen, aber Content leer oder nicht verfügbar." -Level "DEBUG"
            }
        } catch {
            Write_LogEntry -Message "Fehler beim Abrufen der Redirect-Seite $($redirectUrl): $($_)" -Level "ERROR"
            $redirectPage = $null
        }

        # ===== Extract download link robustly (parsed links first, regex fallback) =====
        $downloadLink = $null

        if ($redirectPage) {
            try {
                if ($redirectPage -and $redirectPage.Links) {
                    $linkObj = $redirectPage.Links | Where-Object {
                        ($_.href -and ($_.href -match '\.exe($|\?)')) -or ($_.innerText -and $_.innerText -match '\.exe')
                    } | Select-Object -First 1

                    if ($linkObj) { $downloadLink = $linkObj.href }
                }
            } catch {
                Write_LogEntry -Message "Fehler beim Auslesen von Links aus dem Redirect-Objekt: $($_)" -Level "DEBUG"
            }

            # fallback: search HTML for absolute or protocol-relative .exe URLs
            if (-not $downloadLink -and $redirectPage.Content) {
                $exeMatch = [regex]::Match($redirectPage.Content, '(https?:)?\/\/[^"''\s]+?\.exe', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($exeMatch.Success) {
                    $downloadLink = $exeMatch.Value
                } else {
                    # last resort: any href="...exe" (relative or absolute)
                    $exeMatch = [regex]::Match($redirectPage.Content, 'href\s*=\s*["'']([^"'']+?\.exe(?:\?[^"'']*)?)["'']', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($exeMatch.Success) { $downloadLink = $exeMatch.Groups[1].Value }
                }
            }

            # normalize protocol-relative (//...) and relative links to absolute
            if ($downloadLink) {
                if ($downloadLink -match '^\/\/') {
                    $downloadLink = "https:$downloadLink"
                } elseif ($downloadLink -notmatch '^https?:\/\/') {
                    # try to find a base URI from the response object in a PS-version-safe manner
                    $baseUri = $null
                    try { if ($redirectPage.BaseResponse -and $redirectPage.BaseResponse.ResponseUri) { $baseUri = $redirectPage.BaseResponse.ResponseUri } } catch {}
                    if (-not $baseUri) {
                        try { if ($redirectPage.ResponseUri) { $baseUri = $redirectPage.ResponseUri } } catch {}
                    }
                    if (-not $baseUri) {
                        try { if ($redirectPage.BaseResponse -and $redirectPage.BaseResponse.RequestMessage -and $redirectPage.BaseResponse.RequestMessage.RequestUri) { $baseUri = $redirectPage.BaseResponse.RequestMessage.RequestUri } } catch {}
                    }

                    try {
                        if ($baseUri) {
                            $downloadLink = [System.Uri]::new($baseUri, $downloadLink).AbsoluteUri
                        } else {
                            # as a last resort, prepend https://winscp.net
                            if ($downloadLink -notmatch '^\/') { $downloadLink = "https://winscp.net/$downloadLink" } else { $downloadLink = "https://winscp.net$downloadLink" }
                        }
                    } catch {
                        Write_LogEntry -Message "Konnte relative URL nicht normalisieren: $($downloadLink) - $($_)" -Level "WARNING"
                    }
                }
            }
        }

        Write_LogEntry -Message "Extrahierter Download-Link: $($downloadLink)" -Level "DEBUG"

        if ($downloadLink) {
            # Extract filename and prepare download path
            try {
                $uriObj = [System.Uri]::new($downloadLink)
                $filename = [System.IO.Path]::GetFileName($uriObj.AbsolutePath)
            } catch {
                $filename = "WinSCP-$latestVersion-Setup.exe"
            }
            $downloadPath = Join-Path -Path $InstallationFolder -ChildPath $filename
            Write_LogEntry -Message "Downloade $($downloadLink) nach $($downloadPath)" -Level "INFO"

            # ********** Use System.Net.WebClient as primary downloader (fast), fallback to Invoke-WebRequest if needed **********
            $downloadSucceeded = $false
            $webClient = $null
            try {
                $wcType = [type]::GetType("System.Net.WebClient", $false)
                if ($wcType) {
                    $webClient = New-Object System.Net.WebClient

                    # optional: set common headers to mimic a browser (some servers require user-agent)
                    try { $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT) PowerShell") } catch {}
                    # respect system proxy settings
                    try { $webClient.Proxy = [System.Net.WebRequest]::DefaultWebProxy } catch {}

                    # Download (synchronous, fast)
                    [void](Invoke-DownloadFile -Url $downloadLink -OutFile $downloadPath)
                    $downloadSucceeded = Test-Path -Path $downloadPath

                    if ($downloadSucceeded) {
                        Write_LogEntry -Message "Download abgeschlossen (WebClient): $($downloadPath)" -Level "SUCCESS"
                    } else {
                        Write_LogEntry -Message "WebClient-Download abgeschlossen, aber Datei nicht gefunden: $($downloadPath)" -Level "ERROR"
                    }
                } else {
                    throw "System.Net.WebClient type not available in this runtime."
                }
            } catch {
                Write_LogEntry -Message "WebClient-Download fehlgeschlagen: $($_). Versuche Invoke-WebRequest-Fallback." -Level "WARNING"
                # Fallback to Invoke-WebRequest
                try {
                    Invoke-WebRequestCompat -Uri $downloadLink -OutFile $downloadPath
                    $downloadSucceeded = Test-Path -Path $downloadPath
                    if ($downloadSucceeded) {
                        Write_LogEntry -Message "Download abgeschlossen (Invoke-WebRequest Fallback): $($downloadPath)" -Level "SUCCESS"
                    } else {
                        Write_LogEntry -Message "Invoke-WebRequest finishte ohne Ergebnis: $($downloadPath) nicht gefunden" -Level "ERROR"
                    }
                } catch {
                    Write_LogEntry -Message "Invoke-WebRequest-Fallback ebenfalls fehlgeschlagen: $($_)" -Level "ERROR"
                }
            } finally {
                if ($null -ne $webClient) {
                    try { $webClient.Dispose() } catch {}
                }
            }

            # Check if the file was completely downloaded
            if ($downloadSucceeded -and (Test-Path $downloadPath)) {
                # Remove the old installer (if possible)
                try {
                    Remove-Item -Path $InstallationFileFile.FullName -Force -ErrorAction Stop
                    Write_LogEntry -Message "Alte Installationsdatei entfernt: $($InstallationFileFile.FullName)" -Level "DEBUG"
                } catch {
                    Write_LogEntry -Message "Fehler beim Entfernen der alten Installationsdatei $($InstallationFileFile.FullName): $($_)" -Level "WARNING"
                }

                Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
                Write_LogEntry -Message "$($ProgramName) Update erfolgreich heruntergeladen: $($downloadPath)" -Level "SUCCESS"
            } else {
                Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -foregroundcolor "red"
                Write_LogEntry -Message "Download fehlgeschlagen: $($downloadPath) nicht gefunden nach Download" -Level "ERROR"
            }
        } else {
            Write_LogEntry -Message "Kein Download-Link in der Seite gefunden: $($redirectUrl)" -Level "WARNING"
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
        Write_LogEntry -Message "Kein Update verfügbar: Online $($latestVersion) <= Local $($localVersion)" -Level "INFO"
    }

    Write-Host ""
    Write_LogEntry -Message "Beginne Prüfung installierter Versionen in der Registry" -Level "DEBUG"

    #Check Installed Version / Install if needed
    $InstallationFileFile = Get-ChildItem -Path $installerPath -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($InstallationFileFile) {
        $versionInfo = (Get-Item $InstallationFileFile).VersionInfo
        $localVersion = ($versionInfo.ProductVersion).ToString().Trim()
        Write_LogEntry -Message "Erneut gefundene Installationsdatei: $($InstallationFileFile.FullName); Version: $($localVersion)" -Level "DEBUG"
    } else {
        Write_LogEntry -Message "Keine Installationsdatei gefunden beim zweiten Scan: $($installerPath)" -Level "WARNING"
    }

    $RegistryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    Write_LogEntry -Message "Registry-Pfade für Suche konfiguriert: $($RegistryPaths -join ', ')" -Level "DEBUG"

    $Path = foreach ($RegPath in $RegistryPaths) {
        if (Test-Path $RegPath) {
            Write_LogEntry -Message "Registry-Pfad existiert: $($RegPath)" -Level "DEBUG"
            Get-ChildItem $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "$ProgramName*" }
        } else {
            Write_LogEntry -Message "Registry-Pfad nicht vorhanden: $($RegPath)" -Level "DEBUG"
        }
    }

    if ($null -ne $Path) {
        $installedVersion = $Path.DisplayVersion | Select-Object -First 1
        Write-Host "$ProgramName ist installiert." -foregroundcolor "green"
        Write-Host "    Installierte Version:       $installedVersion" -foregroundcolor "Cyan"
        Write-Host "    Installationsdatei Version: $localVersion" -foregroundcolor "Cyan"
        Write_LogEntry -Message "Gefundene installierte Version aus Registry: $($installedVersion); Datei-Version: $($localVersion)" -Level "INFO"

        try {
            if ([version]$installedVersion -lt [version]$localVersion) {
                Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -foregroundcolor "magenta"
                $Install = $true
                Write_LogEntry -Message "Install erforderlich: Registry $($installedVersion) < Datei $($localVersion)" -Level "INFO"
            } elseif ([version]$installedVersion -eq [version]$localVersion) {
                Write-Host "        Installierte Version ist aktuell." -foregroundcolor "DarkGray"
                $Install = $false
                Write_LogEntry -Message "Install nicht erforderlich: InstalledVersion == LocalVersion ($($localVersion))" -Level "INFO"
            } else {
                $Install = $false
                Write_LogEntry -Message "Install nicht ausgeführt: InstalledVersion ($($installedVersion)) > LocalVersion ($($localVersion))" -Level "WARNING"
            }
        } catch {
            # if version cast fails, fallback to string compare
            Write_LogEntry -Message "Fehler beim Vergleichen der Versionen via [version]: $($_). Fallback-Stringvergleich." -Level "WARNING"
            if (Test-IsVersionGreater -a $localVersion $installedVersion) {
                $Install = $true
            } else {
                $Install = $false
            }
        }
    } else {
        $Install = $false
        Write_LogEntry -Message "Keine Registry-Einträge für $($ProgramName) gefunden" -Level "DEBUG"
    }
    Write-Host ""

    #Install if needed
    if($InstallationFlag){
        Write_LogEntry -Message "Starte externes Installationsskript wegen InstallationFlag: $($InstallationFlag)" -Level "INFO"
        & $PSHostPath `
            -NoLogo -NoProfile -ExecutionPolicy Bypass `
            -File "$Serverip\Daten\Prog\InstallationScripts\Installation\WinScpInstall.ps1" `
            -InstallationFlag
        Write_LogEntry -Message "Externer Aufruf abgeschlossen: WinScpInstall.ps1 mit -InstallationFlag" -Level "DEBUG"
    } elseif($Install -eq $true){
        Write_LogEntry -Message "Starte externes Installationsskript wegen Install=true" -Level "INFO"
        & $PSHostPath `
            -NoLogo -NoProfile -ExecutionPolicy Bypass `
            -File "$Serverip\Daten\Prog\InstallationScripts\Installation\WinScpInstall.ps1"
        Write_LogEntry -Message "Externer Aufruf abgeschlossen: WinScpInstall.ps1" -Level "DEBUG"
    }
    Write-Host ""
} else {
    Write_LogEntry -Message "Kein WinSCP-Installer gefunden im Pfad: $($installerPath)" -Level "WARNING"
}

# === Logger-Footer: automatisch eingefügt ===
if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage "$ProgramName - Script beendet"
} else {
    Write_LogEntry -Message "$ProgramName - Script beendet" -Level "INFO"
}
# === Ende Logger-Footer ===
