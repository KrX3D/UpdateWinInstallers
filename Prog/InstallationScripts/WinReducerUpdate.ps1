param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "WinReducerEX100"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
. (Get-SharedConfigPath -ScriptRoot $PSScriptRoot)   # exposes $NetworkShareDaten

$InstallationFolder  = "$NetworkShareDaten\Customize_Windows\Tools\WinReducerEX100"
$InstallationExe     = "$InstallationFolder\WinReducerEX100_x64.exe"
$destinationPath     = Join-Path $env:USERPROFILE "Desktop\WinReducerEX100"
$destinationFilePath = Join-Path $destinationPath "WinReducerEX100_x64.exe"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder | Exe: $InstallationExe" -Level 'DEBUG'

# Online version from WinReducer website
$downloadPageUrl = "https://www.winreducer.net/storage-plus/e2c12abf-f4f5-4d2b-b9e0-259180c3c010"
$webpageContent  = ""
$webSession      = New-Object Microsoft.PowerShell.Commands.WebRequestSession

try {
    $webResponse = Invoke-WebRequest -Uri $downloadPageUrl -WebSession $webSession -UseBasicParsing -ErrorAction Stop
    $webpageContent = $webResponse.Content
    Write-DeployLog -Message "Storage-Seite abgerufen: $downloadPageUrl" -Level 'DEBUG'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Storage-Seite $downloadPageUrl : $_" -Level 'ERROR'
}

$versionText  = ""
$downloadURL  = ""
$zipFileName  = ""

if ($webpageContent) {
    # Primary strategy for Wix Storage+: call file-sharing APIs directly.
    $storageIdMatch = [regex]::Match($downloadPageUrl, '[0-9a-fA-F-]{36}$')
    if ($storageIdMatch.Success) {
        $storageId = $storageIdMatch.Value
        $headers = @{
            'Content-Type' = 'application/json'
            'Accept'       = 'application/json, text/plain, */*'
            'Origin'       = 'https://www.winreducer.net'
            'Referer'      = $downloadPageUrl
            'x-wix-brand'  = 'wix'
        }

        # Bootstrap Wix visitor session/token endpoint to avoid permission denied on file-sharing API.
        try {
            $null = Invoke-RestMethod -Uri 'https://www.winreducer.net/_api/v1/access-tokens' -Method Get -WebSession $webSession -Headers @{ 'Accept' = 'application/json, text/plain, */*' } -ErrorAction Stop
            Write-DeployLog -Message 'Wix Access-Token Endpoint erfolgreich initialisiert.' -Level 'DEBUG'
        } catch {
            Write-DeployLog -Message "Wix Access-Token Bootstrap fehlgeschlagen: $_" -Level 'WARNING'
        }

        try {
            $xsrfCookie = $webSession.Cookies.GetCookies('https://www.winreducer.net') | Where-Object { $_.Name -eq 'XSRF-TOKEN' } | Select-Object -First 1
            if ($xsrfCookie -and $xsrfCookie.Value) {
                $headers['x-xsrf-token'] = [System.Uri]::UnescapeDataString($xsrfCookie.Value)
                Write-DeployLog -Message 'XSRF-Token aus Cookie gesetzt.' -Level 'DEBUG'
            }
        } catch {
            Write-DeployLog -Message "XSRF-Token konnte nicht gelesen werden: $_" -Level 'WARNING'
        }

        $apiBase = 'https://www.winreducer.net/api/v1/file-sharing'
        $queryBody = @{ filter = @{ parentLibraryItemIds = @($storageId) }; sort = @{ orientation = 'ASC'; sortBy = 'NAME' } } | ConvertTo-Json -Depth 6

        try {
            $queryResponse = Invoke-RestMethod -Uri "$apiBase/library-items/query" -Method Post -WebSession $webSession -Headers $headers -Body $queryBody -ErrorAction Stop
            $items = @($queryResponse.libraryItems)
            Write-DeployLog -Message "Storage API: $($items.Count) Eintraege gefunden." -Level 'DEBUG'

            $x64Items = @($items | Where-Object { $_.name -match '^WinReducer_EX_Series_x64_(\d+(?:\.\d+){1,3})\.zip$' })
            if ($x64Items.Count -gt 0) {
                $candidates = @()
                foreach ($item in $x64Items) {
                    $m = [regex]::Match($item.name, '^WinReducer_EX_Series_x64_(\d+(?:\.\d+){1,3})\.zip$')
                    if ($m.Success) {
                        try {
                            $candidates += [pscustomobject]@{
                                Name    = $item.name
                                ItemId  = $item.id
                                Version = $m.Groups[1].Value
                                Parsed  = [version]$m.Groups[1].Value
                            }
                        } catch {}
                    }
                }

                if ($candidates.Count -gt 0) {
                    $latest = $candidates | Sort-Object Parsed -Descending | Select-Object -First 1
                    $versionText = $latest.Version
                    $zipFileName = $latest.Name

                    $viewBody = @{ actions = @(@{ libraryItemId = $latest.ItemId }) } | ConvertTo-Json -Depth 6
                    $viewResponse = Invoke-RestMethod -Uri "$apiBase/library-items/view-file" -Method Post -WebSession $webSession -Headers $headers -Body $viewBody -ErrorAction Stop
                    $downloadURL = @($viewResponse.urls | Select-Object -ExpandProperty url -ErrorAction SilentlyContinue | Select-Object -First 1)
                    Write-DeployLog -Message "Storage API Download-Link ermittelt fuer Item $($latest.ItemId)." -Level 'DEBUG'
                }
            }
        } catch {
            Write-DeployLog -Message "Storage API-Abfrage fehlgeschlagen: $_" -Level 'WARNING'
        }
    }

    # Fallback strategy: parse rendered/encoded HTML data when API path fails.
    if (-not $zipFileName) {
        $decodedHtml = [System.Net.WebUtility]::HtmlDecode($webpageContent)
        $decodedUri  = [System.Uri]::UnescapeDataString($webpageContent)
        $decodedBoth = [System.Uri]::UnescapeDataString($decodedHtml)
        $searchBlobs = @($webpageContent, $decodedHtml, $decodedUri, $decodedBoth)

        $zipPattern = 'WinReducer_EX_Series_x64_(\d+(?:\.\d+){1,3})\.zip'
        $allVersions = @()
        foreach ($blob in $searchBlobs) {
            $versionMatches = [regex]::Matches($blob, $zipPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $versionMatches) {
                $v = $match.Groups[1].Value
                try { $allVersions += [pscustomobject]@{ Text = $v; Parsed = [version]$v } } catch {}
            }
        }

        if ($allVersions.Count -gt 0) {
            $latest = $allVersions | Sort-Object Parsed -Descending | Select-Object -First 1
            $versionText = $latest.Text
            $zipFileName = "WinReducer_EX_Series_x64_$versionText.zip"
            $downloadURL = "$downloadPageUrl/$zipFileName"
            Write-DeployLog -Message "Fallback-Parsing verwendet: $zipFileName" -Level 'DEBUG'
        }
    }
}

if (-not $downloadURL -and $zipFileName) {
    $downloadURL = "$downloadPageUrl/$zipFileName"
    Write-DeployLog -Message "Keine direkte Download-URL in HTML gefunden, verwende Fallback-URL: $downloadURL" -Level 'WARNING'
}

Write-DeployLog -Message "Online-Version: $versionText" -Level 'INFO'
Write-DeployLog -Message "Online-Datei: $zipFileName" -Level 'DEBUG'
Write-DeployLog -Message "Download-URL: $downloadURL" -Level 'DEBUG'

if (-not $versionText -and $webpageContent) {
    $snippet = $webpageContent.Substring(0, [Math]::Min(800, $webpageContent.Length))
    Write-DeployLog -Message "Keine x64 ZIP-Version gefunden. HTML-Start (800 Zeichen): $snippet" -Level 'WARNING'
}

# Local version (FileVersion from exe)
$localVersion = ""
try {
    $localVersion = (Get-ItemProperty -Path $InstallationExe -ErrorAction Stop).VersionInfo.FileVersion
    Write-DeployLog -Message "Lokale Version: $localVersion" -Level 'DEBUG'
} catch {
    Write-DeployLog -Message "Fehler beim Lesen der lokalen Version: $_" -Level 'WARNING'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $versionText"   -ForegroundColor Cyan
Write-Host ""

# Download and update if newer
$onlineVersion = $null
$localVersionObj = $null
try { if ($versionText) { $onlineVersion = [version]$versionText } } catch {
    Write-DeployLog -Message "Online-Version nicht gueltig: $versionText" -Level 'WARNING'
}
try { if ($localVersion) { $localVersionObj = [version]$localVersion } } catch {
    Write-DeployLog -Message "Lokale Version nicht gueltig: $localVersion" -Level 'WARNING'
}

if ($onlineVersion -and (-not $localVersionObj -or $onlineVersion -gt $localVersionObj) -and $downloadURL) {
    Write-DeployLog -Message "Update verfuegbar: $localVersion -> $versionText" -Level 'INFO'
    $downloadPath = Join-Path $env:TEMP "WinReducerEX100_x64_new.zip"

    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    try {
        Invoke-DownloadFile -Url $downloadURL -OutFile $downloadPath | Out-Null
        Write-DeployLog -Message "Download abgeschlossen: $downloadPath" -Level 'SUCCESS'
    } catch {
        Write-DeployLog -Message "Fehler beim Download: $_" -Level 'ERROR'
    } finally {
        $null = $wc.Dispose()
    }

    if (Test-Path $downloadPath) {
        try { Expand-Archive -Path $downloadPath -DestinationPath $env:TEMP -Force } catch {
            Write-DeployLog -Message "Fehler beim Entpacken: $_" -Level 'ERROR'
        }

        # Backup licence/config before replacing
        $backupDest = Join-Path $env:TEMP "WinReducer_EX_Series_x64\WinReducerEX100\HOME\SOFTWARE"
        try { Move-Item -Path (Join-Path $InstallationFolder "HOME\SOFTWARE\x64")              -Destination $backupDest -Force } catch {}
        try { Move-Item -Path (Join-Path $InstallationFolder "HOME\SOFTWARE\WinReducerEX100.xml") -Destination $backupDest -Force } catch {}

        # Replace installation folder
        try { Remove-Item -Path $InstallationFolder -Recurse -Force } catch {
            Write-DeployLog -Message "Fehler beim Entfernen des alten Installationsordners: $_" -Level 'WARNING'
        }
        $extractedFolder = Join-Path $env:TEMP "WinReducer_EX_Series_x64\WinReducerEX100"
        try {
            Move-Item -Path $extractedFolder -Destination (Split-Path $InstallationFolder -Parent) -Force
            Write-DeployLog -Message "Extrahierter Ordner verschoben nach: $(Split-Path $InstallationFolder -Parent)" -Level 'SUCCESS'
        } catch {
            Write-DeployLog -Message "Fehler beim Verschieben: $_" -Level 'ERROR'
        }

        # Cleanup temp
        Remove-Item -Path (Join-Path $env:TEMP "WinReducer_EX_Series_x64") -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $downloadPath -Force -ErrorAction SilentlyContinue

        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
        Write-DeployLog -Message "$ProgramName aktualisiert." -Level 'SUCCESS'
    } else {
        Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
        Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
    }
} elseif (-not $downloadURL -and $versionText) {
    Write-Host "Online Version erkannt, aber Download-Link nicht gefunden. $ProgramName wurde nicht aktualisiert." -ForegroundColor Yellow
    Write-DeployLog -Message "Version erkannt ($versionText), aber kein Download-Link gefunden." -Level 'WARNING'
} else {
    Write-Host "Kein Online Update verfuegbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
    Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
}

Write-Host ""

# Re-read local version after potential update
try {
    $localVersion = (Get-ItemProperty -Path $InstallationExe -ErrorAction Stop).VersionInfo.FileVersion
} catch {
    $localVersion = ""
}

# Installed version = file on Desktop
$installedVersion = $null
if (Test-Path $destinationFilePath) {
    try { $installedVersion = (Get-ItemProperty -Path $destinationFilePath).VersionInfo.FileVersion } catch {}
}

$Install = $false
if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Desktop-Version: $installedVersion | Lokal: $localVersion" -Level 'INFO'
    try { $Install = [version]$installedVersion -lt [version]$localVersion } catch { $Install = $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht auf Desktop installiert." -Level 'INFO'
}

Write-Host ""

# Copy to desktop if needed
if ($Install -or $InstallationFlag) {
    Write-Host "WinReducerEX100 wird kopiert" -ForegroundColor Cyan
    Write-DeployLog -Message "Starte Kopiervorgang: $InstallationFolder -> $destinationPath" -Level 'INFO'
    if (Test-Path $InstallationExe) {
        try {
            Copy-Item -Path $InstallationFolder -Destination $destinationPath -Recurse -Force
            Write-DeployLog -Message "Kopiervorgang erfolgreich." -Level 'SUCCESS'
        } catch {
            Write-DeployLog -Message "Fehler beim Kopieren: $_" -Level 'ERROR'
        }
    } else {
        Write-DeployLog -Message "Quellpfad nicht gefunden: $InstallationExe" -Level 'ERROR'
    }
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
