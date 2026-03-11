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

try {
    $webpageContent = (Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing -ErrorAction Stop).Content
    Write-DeployLog -Message "Storage-Seite abgerufen: $downloadPageUrl" -Level 'DEBUG'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Storage-Seite $downloadPageUrl : $_" -Level 'ERROR'
}

$versionText  = ""
$downloadURL  = ""
$zipFileName  = ""

if ($webpageContent) {
    $decodedHtml = [System.Net.WebUtility]::HtmlDecode($webpageContent)
    $decodedUri  = [System.Uri]::UnescapeDataString($webpageContent)
    $decodedBoth = [System.Uri]::UnescapeDataString($decodedHtml)
    $regexUnescaped = [regex]::Unescape($webpageContent)
    $searchBlobs = @($webpageContent, $decodedHtml, $decodedUri, $decodedBoth, $regexUnescaped)

    # Wix/plugin pages may keep file listings behind secondary JSON endpoints.
    $extraDataUrls = @(
        "$downloadPageUrl?format=json",
        "$downloadPageUrl?ajax=true",
        "$downloadPageUrl?view=ajax",
        "$downloadPageUrl/index.json",
        "$downloadPageUrl/data.json",
        "$downloadPageUrl/files.json"
    )
    foreach ($extraUrl in $extraDataUrls) {
        try {
            $extraContent = (Invoke-WebRequest -Uri $extraUrl -UseBasicParsing -ErrorAction Stop).Content
            if ($extraContent) {
                $searchBlobs += $extraContent
                $searchBlobs += [System.Net.WebUtility]::HtmlDecode($extraContent)
                $searchBlobs += [System.Uri]::UnescapeDataString($extraContent)
                $searchBlobs += [regex]::Unescape($extraContent)
                Write-DeployLog -Message "Zusatzdaten geladen: $extraUrl" -Level 'DEBUG'
            }
        } catch {}
    }

    $zipPattern = 'WinReducer_EX_Series_x64_(\d+(?:\\?\.\d+){1,3})\\?\.zip'
    $allVersions = @()
    foreach ($blob in $searchBlobs) {
        $versionMatches = [regex]::Matches($blob, $zipPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $versionMatches) {
            $v = $match.Groups[1].Value -replace '\\.', '.'
            try { $allVersions += [pscustomobject]@{ Text = $v; Parsed = [version]$v } } catch {}
        }

        # Fallback for encoded/escaped filename forms (e.g. %5F, \u005f, \\.)
        $normalizedBlob = $blob
        $normalizedBlob = $normalizedBlob -replace '(?i)%5f', '_'
        $normalizedBlob = $normalizedBlob -replace '(?i)\\u005f', '_'
        $normalizedBlob = $normalizedBlob -replace '\\.', '.'
        $normalizedBlob = $normalizedBlob -replace '\\/', '/'

        $tokenMatches = [regex]::Matches($normalizedBlob, 'WinReducer[^"''\s<>]{0,120}x64[^"''\s<>]{0,120}\.zip', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($token in $tokenMatches) {
            $vm = [regex]::Match($token.Value, 'x64[_-](\d+(?:\.\d+){1,3})\.zip', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($vm.Success) {
                $v = $vm.Groups[1].Value
                try { $allVersions += [pscustomobject]@{ Text = $v; Parsed = [version]$v } } catch {}
            }
        }
    }

    if ($allVersions.Count -gt 0) {
        $latest = $allVersions | Sort-Object Parsed -Descending | Select-Object -First 1
        $versionText = $latest.Text
        $zipFileName = "WinReducer_EX_Series_x64_$versionText.zip"
    }

    if ($zipFileName) {
        $escapedFileName = [regex]::Escape($zipFileName)
        $encodedFileName = [regex]::Escape([System.Uri]::EscapeDataString($zipFileName))
        $urlPatterns = @(
            ('https?:\/\/[^"''\s]+{0}(?:\?[^"''\s]*)?' -f $escapedFileName),
            ('https?://[^"''\s]+{0}(?:\?[^"''\s]*)?' -f $escapedFileName),
            ('https?:\/\/[^"''\s]+{0}(?:\?[^"''\s]*)?' -f $encodedFileName),
            ('https?://[^"''\s]+{0}(?:\?[^"''\s]*)?' -f $encodedFileName),
            ('href="([^"]*{0}[^"]*)"' -f $escapedFileName),
            ('href="([^"]*{0}[^"]*)"' -f $encodedFileName),
            ('"url"\s*:\s*"([^"]*{0}[^"]*)"' -f $escapedFileName)
        )

        foreach ($blob in $searchBlobs) {
            foreach ($pattern in $urlPatterns) {
                $urlMatch = [regex]::Match($blob, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($urlMatch.Success) {
                    $candidate = if ($urlMatch.Groups.Count -gt 1 -and $urlMatch.Groups[1].Value) { $urlMatch.Groups[1].Value } else { $urlMatch.Value }
                    if ($candidate) {
                        $candidate = $candidate.Replace('\\/', '/')
                        if ($candidate -match '^https?://') {
                            $downloadURL = $candidate
                        } elseif ($candidate.StartsWith('/')) {
                            $downloadURL = "https://www.winreducer.net$candidate"
                        }
                    }
                }

                if ($downloadURL) { break }
            }
            if ($downloadURL) { break }
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
