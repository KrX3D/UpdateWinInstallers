param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "NVIDIA Grafiktreiber"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
. (Get-SharedConfigPath -ScriptRoot $PSScriptRoot)   # exposes $NetworkShareDaten

$InstallationFolder = "$NetworkShareDaten\Treiber\AMD_PC"
$installScript      = "$Serverip\Daten\Prog\InstallationScripts\Installation\TreiberAmdPcNvideaInstall.ps1"

Write-DeployLog -Message "InstallationFolder: $InstallationFolder" -Level 'INFO'

# ── PC identity check ──────────────────────────────────────────────────────────
$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
try {
    $PCName = & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
    Write-DeployLog -Message "PC-Name: $PCName" -Level 'INFO'
} catch {
    Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
    Write-DeployLog -Message "Fehler beim PC-Ermittlungs-Script: $_" -Level 'ERROR'
    Exit
}

Write-Host "PC Name: $PCName"

if ($PCName -eq "KrX-AMD-PC") {
    Write-DeployLog -Message "Zielsystem erkannt. Starte NVIDIA-Workflow." -Level 'INFO'

    # Folder pattern: versioned extract folder ending with NSD driver suffix
    $folderPattern  = "desktop-win10-win11-64bit-international-nsd-dch-whql$"
    $cfgVerPattern  = '<setup title="\$\{\{ProductTitle\}\}" version="([\d.]+)"'

    # ── Fetch latest available version from NVIDIA API ─────────────────────────
    function Get-NvidiaVersionAndUrl {
        param([string]$CurrentLocalVersion)
        try {
            $params = @{
                func="DriverManualLookup"; psid="120"; pfid="929"; osID="57"
                languageCode="1033"; beta="0"; isWHQL="1"; dltype="-1"
                dch="1"; upCRD="0"; qnf="0"; sort1="0"; numberOfResults="10"
            }
            $qs       = ($params.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
            $response = Invoke-RestMethod -Uri "https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php?$qs" -Method Get -UseBasicParsing -ErrorAction Stop

            if (-not $response.IDS -or $response.IDS.Count -eq 0) {
                Write-DeployLog -Message "NVIDIA API: keine Treiber zurückgegeben." -Level 'ERROR'
                return $null
            }

            $newer = @()
            foreach ($drv in $response.IDS) {
                $ver = $drv.downloadInfo.Version
                if ($CurrentLocalVersion -and [version]$ver -gt [version]$CurrentLocalVersion) { $newer += $drv }
            }

            if ($newer.Count -eq 0) {
                Write-DeployLog -Message "Keine neueren Treiber als $CurrentLocalVersion verfügbar." -Level 'INFO'
                Write-Host "Keine neueren Treiber als $CurrentLocalVersion verfügbar." -ForegroundColor Green
                return $null
            }

            Write-Host ""
            Write-Host "Verfügbare Updates (aktuell lokal: $CurrentLocalVersion):" -ForegroundColor Cyan
            for ($i = 0; $i -lt $newer.Count; $i++) {
                Write-Host "  [$($i+1)] Version: $($newer[$i].downloadInfo.Version) | $($newer[$i].downloadInfo.ReleaseDateTime)" -ForegroundColor Cyan
            }
            Write-Host ""

            foreach ($drv in $newer) {
                $ver = $drv.downloadInfo.Version
                $url = "https://de.download.nvidia.com/Windows/$ver/$ver-desktop-win10-win11-64bit-international-nsd-dch-whql.exe"
                Write-Host "  -> Prüfe Verfügbarkeit von Version $ver..." -ForegroundColor Yellow
                try {
                    $head = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
                    if ($head.StatusCode -eq 200) {
                        Write-Host "  [OK] Version $ver ist verfügbar!" -ForegroundColor Green; Write-Host ""
                        Write-DeployLog -Message "NVIDIA Version $ver verfügbar: $url" -Level 'SUCCESS'
                        return @{ VersionNumber=$ver; DownloadLink=$url }
                    }
                } catch {
                    Write-Host "  [X] Version $ver noch nicht verfügbar." -ForegroundColor Red
                    Write-DeployLog -Message "Version $ver nicht erreichbar: $_" -Level 'WARNING'
                }
            }

            Write-Host "Keine neueren Treiber zum Download verfügbar." -ForegroundColor Yellow
            Write-DeployLog -Message "Kein verfügbarer Download für neuere NVIDIA-Versionen." -Level 'WARNING'
            return $null
        } catch {
            Write-DeployLog -Message "NVIDIA API Fehler: $_" -Level 'ERROR'
            return $null
        }
    }

    function ExtractExeFile {
        $sz = "C:\Program Files\7-Zip\7z.exe"
        if (-not (Test-Path $sz)) { return }
        $base   = [System.IO.Path]::GetFileNameWithoutExtension($downloadFileName)
        $target = Join-Path $InstallationFolder $base
        if (-not (Test-Path $target)) { New-Item -Path $target -ItemType Directory | Out-Null }
        Write-Host "Die Datei wird extrahiert." -ForegroundColor Yellow
        & "$sz" x "$downloadPath" -o"$target" -y *> $null
    }

    # ── Find local driver folder ───────────────────────────────────────────────
    $matchingFolder = Get-ChildItem -Path $InstallationFolder -Directory | Where-Object { $_.Name -match $folderPattern } | Select-Object -First 1

    if ($matchingFolder) {
        Write-DeployLog -Message "Lokaler Treiber-Ordner: $($matchingFolder.FullName)" -Level 'INFO'

        $cfgPath = Join-Path $matchingFolder.FullName "setup.cfg"
        $localVersion = $null
        if (Test-Path $cfgPath) {
            $cfgContent   = Get-Content -Path $cfgPath -Raw
            $m            = [regex]::Match($cfgContent, $cfgVerPattern)
            if ($m.Success) { $localVersion = $m.Groups[1].Value }
            Write-DeployLog -Message "Lokale Version aus setup.cfg: $localVersion" -Level 'DEBUG'
        } else {
            Write-DeployLog -Message "setup.cfg nicht gefunden: $cfgPath" -Level 'WARNING'
        }

        $result = Get-NvidiaVersionAndUrl -CurrentLocalVersion $localVersion
        if ($result) {
            $onlineVersion = $result.VersionNumber
            $downloadUrl   = "https://de.download.nvidia.com/Windows/$onlineVersion/$onlineVersion-desktop-win10-win11-64bit-international-nsd-dch-whql.exe"

            Write-Host ""
            Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
            Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
            Write-Host ""

            if ([version]$localVersion -ne [version]$onlineVersion) {
                $downloadFileName = [System.IO.Path]::GetFileName($downloadUrl)
                $downloadPath     = Join-Path $InstallationFolder $downloadFileName

                Import-Module BitsTransfer -ErrorAction SilentlyContinue
                if (Get-Module -Name BitsTransfer -ListAvailable) {
                    Start-BitsTransfer -Source $downloadUrl -Destination $downloadPath
                } else {
                    Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath | Out-Null
                }

                if (Test-Path $downloadPath) {
                    ExtractExeFile
                    try {
                        Remove-Item -Path $downloadPath -Force
                        Remove-Item -Path $matchingFolder -Force -Recurse
                        Write-DeployLog -Message "Alte Dateien entfernt." -Level 'INFO'
                    } catch {
                        Write-DeployLog -Message "Fehler beim Entfernen alter Dateien: $_" -Level 'ERROR'
                    }
                    Write-Host ""; Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
                    Write-DeployLog -Message "$ProgramName aktualisiert auf $onlineVersion." -Level 'SUCCESS'
                } else {
                    Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
                    Write-DeployLog -Message "Download fehlgeschlagen." -Level 'ERROR'
                }
            } else {
                Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
                Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
            }
        } else {
            Write-Host "Kein Treiber-Link gefunden." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Kein passender Ordner in $InstallationFolder gefunden." -ForegroundColor Red
        Write-DeployLog -Message "Kein passender Ordner gefunden." -Level 'ERROR'
    }

    Write-Host ""

    # ── Re-evaluate local ──────────────────────────────────────────────────────
    $matchingFolder = Get-ChildItem -Path $InstallationFolder -Directory | Where-Object { $_.Name -match $folderPattern } | Select-Object -First 1
    $localVersion   = $null
    if ($matchingFolder) {
        $cfgPath = Join-Path $matchingFolder.FullName "setup.cfg"
        if (Test-Path $cfgPath) {
            $m = [regex]::Match((Get-Content $cfgPath -Raw), $cfgVerPattern)
            if ($m.Success) { $localVersion = $m.Groups[1].Value }
        }
    }

    # ── Installed version (registry) ───────────────────────────────────────────
    $installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
    $installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
    $Install          = $false

    if ($installedVersion) {
        Write-Host "$ProgramName ist installiert." -ForegroundColor Green
        Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
        Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
        Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'
        try { $Install = [version]$installedVersion -lt [version]$localVersion } catch { $Install = $false }
        if ($Install) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        } else {
            Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        }
    } else {
        Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
        Write-DeployLog -Message "$ProgramName nicht in Registry." -Level 'INFO'
    }

    Write-Host ""

    if ($InstallationFlag) {
        Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
    } elseif ($Install) {
        Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
    }

    Write-Host ""

} else {
    Write-Host ""
    Write-Host "        Treiber sind NICHT für dieses System geeignet." -ForegroundColor Blue
    Write-Host ""
    Write-DeployLog -Message "System $PCName ist nicht Zielsystem. Keine Aktionen." -Level 'INFO'
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
