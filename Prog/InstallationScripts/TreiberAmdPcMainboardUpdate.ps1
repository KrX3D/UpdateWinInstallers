param(
    [switch]$InstallationFlag = $false,

    [Parameter(Mandatory=$false)]
    [string]$Model = "PRIME X670-P WiFi",

    [Parameter(Mandatory=$false)]
    [ValidateSet("Windows 10 64-bit", "Windows 11 64-bit")]
    [string]$OS = "Windows 10 64-bit",

    [Parameter(Mandatory=$false)]
    [bool]$ShowBrowser = $false
)

$ProgramName = "AMD PC Treiber"
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
Write-DeployLog -Message "InstallationFolder: $InstallationFolder | Model: $Model | OS: $OS | ShowBrowser: $ShowBrowser" -Level 'INFO'

# ── PC identity check ──────────────────────────────────────────────────────────
$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
try {
    $PCName = & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
    Write-DeployLog -Message "PC-Name: $PCName" -Level 'INFO'
} catch {
    Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
    Write-DeployLog -Message "Fehler beim PC-Ermittlungs-Script: $_" -Level 'ERROR'
    Pause; Exit
}

Write-Host "PC Name: $PCName"

if ($PCName -eq "KrX-AMD-PC") {
    Write-DeployLog -Message "Zielsystem erkannt. Starte Treiber-Workflow." -Level 'INFO'

    # ── Node.js / Puppeteer check ──────────────────────────────────────────────
    $TempFolder = [System.IO.Path]::GetTempPath()
    $JsFile     = Join-Path $TempFolder "asus_driver_downloader.js"

    if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
        Write-DeployLog -Message "Node.js nicht installiert – installiere via NodeUpdate.ps1" -Level 'WARNING'
        Write-Host "Node.js is not installed. Installing Node.js." -ForegroundColor Yellow
        & $PSHostPath -ExecutionPolicy Bypass -NoLogo -NoProfile -File "$Serverip\Daten\Prog\InstallationScripts\NodeUpdate.ps1" -InstallationFlag
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }

    # ── Puppeteer install / update ─────────────────────────────────────────────
    $npmWorkingFolder = $env:USERPROFILE
    Push-Location $npmWorkingFolder

    $puppeteerInstalled = $false
    try { npm list puppeteer --depth=0 | Out-Null; $puppeteerInstalled = $true } catch { $puppeteerInstalled = $false }

    if (-not $puppeteerInstalled) {
        Write-Host "Puppeteer not found – installing..." -ForegroundColor Yellow
        npm install puppeteer
    } else {
        $outdated = npm outdated puppeteer --parseable
        if ($outdated) {
            Write-Host "Puppeteer is out-of-date – updating..." -ForegroundColor Yellow
            npm update puppeteer
        } else {
            Write-Host "Puppeteer is already up-to-date." -ForegroundColor Green
        }
    }
    Pop-Location

    # ── JavaScript (Puppeteer scraper) ─────────────────────────────────────────
$JsCode = @'
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

const args = process.argv.slice(2);
const model = args[0] || "PRIME X670-P WiFi";
const os = args[1] || "Windows 10 64-bit";
const showBrowser = args[2] === "true";

const osIdMap = {
  "Windows 10 64-bit": "45",
  "Windows 11 64-bit": "52"
};

async function extractASUSDrivers(model, os) {
  const browser = await puppeteer.launch({
    headless: !showBrowser,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    defaultViewport: { width: 1920, height: 1080 }
  });
  try {
    const page = await browser.newPage();
    const osId = osIdMap[os] || "45";
    const formattedModel = model.replace(/ /g, '_');
    const modelLower = formattedModel.toLowerCase();
    const apiUrl = `https://www.asus.com/support/api/product.asmx/GetPDDrivers?website=de&model=${modelLower}&pdhashedid=&cpu=&osid=${osId}&pdid=99999&siteID=www&sitelang=`;
    const apiResponse = await page.goto(apiUrl, { waitUntil: 'networkidle2', timeout: 60000 });
    const apiText = await apiResponse.text();
    const osForFilename = os.replace(/\s+/g, '_');
    fs.writeFileSync(path.join(require('os').tmpdir(), `asus_api_raw_${osForFilename}.txt`), apiText);
    try {
      const apiJson = JSON.parse(apiText);
      fs.writeFileSync(path.join(require('os').tmpdir(), `asus_api_${osForFilename}.json`), JSON.stringify(apiJson, null, 2));
      const drivers = processApiResponse(apiJson);
      const outputPath = path.join(require('os').tmpdir(), `ASUS_${model.replace(/\s+/g, '_')}_${osForFilename}_drivers.json`);
      fs.writeFileSync(outputPath, JSON.stringify(drivers, null, 2));
      if (drivers.length === 0) { await fallbackWebpageExtraction(page, model, os, formattedModel); }
      return drivers;
    } catch (e) {
      await fallbackWebpageExtraction(page, model, os, formattedModel);
      return [];
    }
  } catch (error) {
    console.error("Script failed:", error);
    fs.writeFileSync(path.join(require('os').tmpdir(), 'asus_error.txt'), String(error.stack || error), 'utf8');
    return [];
  } finally {
    await browser.close();
  }
}

async function fallbackWebpageExtraction(page, model, os, formattedModel) {
  try {
    const url = `https://www.asus.com/de/supportonly/${formattedModel}/helpdesk_download`;
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });
    try {
      await page.waitForSelector('.SortSelect__selectContent__1u-8d', { timeout: 15000 });
      await page.click('.SortSelect__selectContent__1u-8d');
      await new Promise(resolve => setTimeout(resolve, 2000));
      const currentOS = await page.evaluate(() => {
        const el = document.querySelector('.SortSelect__selectContent__1u-8d');
        return el ? el.textContent.trim() : null;
      });
      if (currentOS !== os) {
        let winOption = await page.$(`div[aria-label="${os}"]`) || await page.$(`div[aria-label*="${os}"]`);
        if (winOption) { await winOption.click(); await new Promise(resolve => setTimeout(resolve, 5000)); }
      }
    } catch {}
    const osForFilename = os.replace(/\s+/g, '_');
    await page.screenshot({ path: path.join(require('os').tmpdir(), `asus_page_${osForFilename}.png`), fullPage: true });
    await autoScroll(page);
    await new Promise(resolve => setTimeout(resolve, 5000));
    const pageDrivers = await extractDriversFromPage(page);
    if (pageDrivers.length > 0) {
      const fallbackPath = path.join(require('os').tmpdir(), `ASUS_${model.replace(/\s+/g, '_')}_${osForFilename}_drivers_fallback.json`);
      fs.writeFileSync(fallbackPath, JSON.stringify(pageDrivers, null, 2));
    }
    const html = await page.content();
    fs.writeFileSync(path.join(require('os').tmpdir(), `asus_page_${osForFilename}_full.html`), html);
  } catch (err) { console.error("Fallback failed:", err.message); }
}

function processApiResponse(json) {
  const drivers = [];
  const vReList = [
    /Brand-new Armoury Crate![\s\S]*?Version:[\s\S]*?v([\d.]+)/,
    /Version:\s*-?\s*v([\d.]+)/,
    /Version\s*v([\d.]+)/
  ];
  if (!json.Result || !Array.isArray(json.Result.Obj)) return drivers;
  json.Result.Obj.forEach(category => {
    (category.Files || []).forEach(file => {
      const desc = file.Description || "";
      let svcVer = null;
      for (const re of vReList) { const m = desc.match(re); if (m && m[1]) { svcVer = m[1]; break; } }
      drivers.push({
        category: category.Name || "Unknown", name: file.Title || "Unknown",
        version: file.Version || "Unknown", serviceVersion: svcVer,
        fileSize: file.FileSize || "Unknown", releaseDate: file.ReleaseDate || "Unknown",
        downloadUrl: file.DownloadUrl?.Global || "Unknown", id: file.Id || "Unknown"
      });
    });
  });
  return drivers;
}

async function extractDriversFromPage(page) {
  try {
    return await page.evaluate(() => {
      const drivers = [];
      const boxes = document.querySelectorAll('.productSupportDriverBIOSBox, .ProductSupportDriverBIOS__productSupportDriverBIOSBox__ihsCB');
      boxes.forEach(box => {
        const titleEl = box.querySelector('.ProductSupportDriverBIOS__fileTitle__GE44d');
        const title = titleEl ? titleEl.textContent.trim() : "Unknown";
        const infoContainer = box.querySelector('.ProductSupportDriverBIOS__fileInfo__2c5GN');
        let version = "Unknown";
        const versionEl = infoContainer ? infoContainer.querySelector('div:first-child') : null;
        if (versionEl && versionEl.textContent.includes('Version')) { version = versionEl.textContent.replace('Version','').trim(); }
        let fileSize = "Unknown", releaseDate = "Unknown", category = "Unknown", downloadUrl = "Unknown";
        const sizeEl = infoContainer ? infoContainer.querySelector('.ProductSupportDriverBIOS__fileSize__eticu') : null;
        if (sizeEl) { fileSize = sizeEl.textContent.trim(); }
        const dateEl = infoContainer ? infoContainer.querySelector('.ProductSupportDriverBIOS__releaseDate__3o309') : null;
        if (dateEl) { releaseDate = dateEl.textContent.trim(); }
        const parentSection = box.closest('[class*="type-"]');
        if (parentSection) { const catEl = parentSection.querySelector('h2, .title'); if (catEl) { category = catEl.textContent.trim(); } }
        const btn = box.querySelector('.ProductSupportDriverBIOS__downloadBtn__204JI, .SolidButton__btn__1NmTw');
        if (btn) { downloadUrl = btn.getAttribute('href') || btn.getAttribute('data-href') || btn.getAttribute('data-url') || btn.getAttribute('data-downloadurl') || "Unknown"; }
        drivers.push({ category, name: title, version, fileSize, releaseDate, downloadUrl, extractionMethod: "page_content" });
      });
      return drivers;
    });
  } catch (err) { console.error("Error extracting:", err.message); return []; }
}

async function autoScroll(page) {
  await page.evaluate(async () => {
    await new Promise((resolve) => {
      let totalHeight = 0; const distance = 100;
      const timer = setInterval(() => {
        window.scrollBy(0, distance); totalHeight += distance;
        if (totalHeight >= document.body.scrollHeight) { clearInterval(timer); resolve(); }
      }, 100);
    });
  });
}

extractASUSDrivers(model, os).then(() => {});
'@

    Set-Content -Path $JsFile -Value $JsCode -Encoding UTF8

    Write-Host "Model: $Model" -ForegroundColor Yellow
    Write-Host "OS: $OS" -ForegroundColor Yellow
    if ($ShowBrowser) {
        Write-Host "NOTE: A browser window will open – do not close it until the script finishes." -ForegroundColor Yellow
    } else {
        Write-Host "Running in headless mode." -ForegroundColor Yellow
    }
    Write-Host ""

    node $JsFile "$Model" "$OS" "$ShowBrowser"
    Write-DeployLog -Message "Node.js Puppeteer-Script beendet." -Level 'INFO'

    # ── Parse results ──────────────────────────────────────────────────────────
    $OSForFilename = $OS.Replace(" ", "_")
    $ResultFile    = Get-ChildItem -Path $TempFolder -Filter "ASUS_$($Model.Replace(' ','_'))_${OSForFilename}_drivers.json" | Select-Object -First 1

    if ($ResultFile) {
        $DriversJson = Get-Content $ResultFile.FullName -Raw | ConvertFrom-Json
        Write-Host "Found $($DriversJson.Count) drivers for $Model ($OS)" -ForegroundColor Green
        Write-DeployLog -Message "Treiber gefunden: $($DriversJson.Count)" -Level 'INFO'
    } else {
        $FallbackFile = Get-ChildItem -Path $TempFolder -Filter "ASUS_$($Model.Replace(' ','_'))_${OSForFilename}_drivers_fallback.json" | Select-Object -First 1
        if ($FallbackFile) {
            $DriversJson = Get-Content $FallbackFile.FullName -Raw | ConvertFrom-Json
            Write-Host "Found $($DriversJson.Count) drivers (fallback) for $Model ($OS)" -ForegroundColor Yellow
            $ResultFile = $FallbackFile
        } else {
            Write-Host "No driver information file found. Check for errors." -ForegroundColor Red
            $ErrorFile = Join-Path $TempFolder "asus_error.txt"
            if (Test-Path $ErrorFile) { Get-Content $ErrorFile | Write-Host -ForegroundColor Red }
            Write-DeployLog -Message "Keine Treiber-Informationsdatei gefunden." -Level 'ERROR'
        }
    }

    Remove-Item -Path $JsFile -Force -ErrorAction SilentlyContinue
    Get-ChildItem -Path $TempFolder -Filter "asus_api_*.*" | Remove-Item -Force -ErrorAction SilentlyContinue

    $outputPath = (Get-ChildItem -Path $TempFolder -Filter "ASUS_*_drivers.json" | Select-Object -First 1)?.FullName
    if (-not $outputPath -or -not (Test-Path $outputPath)) {
        Write-DeployLog -Message "ASUS download JSON nicht gefunden." -Level 'ERROR'
        Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
        return
    }

    $allDrivers = Get-Content -Path $outputPath -Raw | ConvertFrom-Json
    Remove-Item -Path $outputPath -Force -ErrorAction SilentlyContinue

    # ── Filter / deduplicate ───────────────────────────────────────────────────
    $excludeKeywords = @("AI Suite","WinRAR","MY ASUS","CPU-Z","Armoury Crate Uninstall Tool","for Ryzen","RAIDXpert2","Aura Creator Installer")
    $filtered = $allDrivers | Where-Object {
        $item = $_
        $exclude = $false
        foreach ($kw in $excludeKeywords) { if ($item.name -like "*$kw*") { $exclude = $true; break } }
        if ($item.name -eq "Armoury Crate & Aura Creator Installer") { $exclude = $false }
        -not $exclude
    }

    function Get-BaseName { param($name)
        if ($name -match '^(.*?)\s+[Vv]\d') { return $Matches[1].Trim() }
        if ($name -match '^(.*?)\s+for\s+Windows') { return $Matches[1].Trim() }
        return $name.Trim()
    }

    $latestPerGroup = $filtered |
        Select-Object *, @{ Name='BaseName'; Expression={ Get-BaseName $_.name } } |
        Group-Object -Property BaseName |
        ForEach-Object {
            $_.Group |
                Sort-Object @{ Expression={ [datetime]($_.releaseDate) } }, @{ Expression={ [version]($_.version) } } -Descending |
                Select-Object -First 1
        }

    $downloadData = $latestPerGroup | ForEach-Object {
        [PSCustomObject]@{
            Label          = $_.name
            DriverName     = $_.BaseName
            Version        = $_.version
            ServiceVersion = $_.serviceVersion
            DownloadLink   = $_.downloadUrl
            releaseDate    = $_.releaseDate
        }
    }

    $driverPatterns = @{
        "Graphics Driver"  = "AMD\s*Graphics|VGA"
        "RAID Driver"      = "RAID"
        "Bluetooth Driver" = "Bluetooth"
        "LAN Driver"       = "LAN|Ethernet"
        "Wi-Fi Driver"     = "Wi[-]?Fi|Wireless"
        "Audio Driver"     = "Audio"
        "Armoury Crate"    = "Armoury\s*Crate"
        "Chipset Driver"   = "AMD\s*Chipset"
    }
    $downloadData | ForEach-Object {
        $grp = "Other"
        foreach ($key in $driverPatterns.Keys) { if ($_.DriverName -match $driverPatterns[$key]) { $grp = $key; break } }
        $_ | Add-Member -MemberType NoteProperty -Name DriverGroup -Value $grp -Force
    }

    $latestDrivers = $downloadData |
        Group-Object -Property DriverGroup |
        ForEach-Object {
            $_.Group | Sort-Object @{Expression={$_.releaseDate}}, @{Expression={$_.Version}} -Descending | Select-Object -First 1
        }

    Write-DeployLog -Message "Online-Treiber-Gruppen: $($latestDrivers.Count)" -Level 'INFO'

    # ── Gather local driver versions ───────────────────────────────────────────
    Function GatherLocalDriverInfos {
        $localdriverPatterns = @{
            "Graphics Driver"  = @{ Path="$InstallationFolder\DRV_VGA_AMD_*";            File="AsusSetup.ini"; Section="InstallInfo"; Key="Version"; Delimiter="false" }
            "RAID Driver"      = @{ Path="$InstallationFolder\DRV_RAID_AMD_RAID_Driver_*"; File="*\*.inf";      Section="Version";     Key="DriverVer"; Delimiter="," }
            "Bluetooth Driver" = @{ Path="$InstallationFolder\DRV_BT_RTK_8852BE_*";       File="AsusSetup.ini"; Section="InstallInfo"; Key="Version"; Delimiter="false" }
            "LAN Driver"       = @{ Path="$InstallationFolder\DRV_LAN_*";                File="AsusSetup.ini"; Section="InstallInfo"; Key="Version"; Delimiter="false" }
            "Wi-Fi Driver"     = @{ Path="$InstallationFolder\DRV_WiFi_RTK_8852BE_*";    File="AsusSetup.ini"; Section="InstallInfo"; Key="Version"; Delimiter="false" }
            "Audio Driver"     = @{ Path="$InstallationFolder\DRV_Audio_RTK_*";          File="AsusSetup.ini"; Section="InstallInfo"; Key="Version"; Delimiter="false" }
            "Armoury Crate"    = @{ Path="$InstallationFolder\ArmouryCrateInstaller_*";   File="ArmouryCrateInstaller.exe"; Section=""; Key=""; Delimiter="FileVersion" }
            "Chipset Driver"   = @{ Path="$InstallationFolder\DRV_Chipset_*";            File="AsusSetup.ini"; Section="InstallInfo"; Key="Version"; Delimiter="false" }
        }
        $driverVersions = @()
        foreach ($driver in $localdriverPatterns.GetEnumerator()) {
            $cfg          = $driver.Value
            $fullPattern  = [System.IO.Path]::Combine($cfg.Path, $cfg.File)
            $basePath     = Split-Path -Path $fullPattern
            $wildcardDir  = (Split-Path -Leaf (Split-Path -Path $fullPattern -Parent))
            $fileName     = Split-Path -Leaf $fullPattern
            $cleanBase    = if ($basePath -match '\\\*$') { $basePath -replace '\\\*$','' } else { $basePath }
            $matchedDirs  = Get-ChildItem -Path $cleanBase -Directory -Filter $wildcardDir -ErrorAction SilentlyContinue
            $cleanedDirs  = $matchedDirs
            $matchedFiles = @()
            foreach ($dir in $matchedDirs) {
                $matchedFiles += Get-ChildItem -Path $dir.FullName -Filter $fileName -ErrorAction SilentlyContinue
            }
            if (-not $matchedFiles) { continue }
            $versions = @()
            foreach ($file in $matchedFiles) {
                if ($file.Extension -in @('.ini','.inf')) {
                    $iniContent = Get-Content -Path $file.FullName
                    $inSection = $false
                    foreach ($line in $iniContent) {
                        if ($line -match "^\[$($cfg.Section)\]") { $inSection = $true; continue }
                        if ($inSection -and $line -match "^\[") { break }
                        if ($inSection -and $line -match "^\s*$($cfg.Key)\s*=\s*(.+)") {
                            $val = $matches[1]
                            if ($cfg.Delimiter -and $cfg.Delimiter -ne "false") { $val = ($val -split $cfg.Delimiter)[-1] }
                            $versions += [string]$val
                        }
                    }
                } elseif ($file.Extension -eq '.exe' -and $cfg.Delimiter -eq "FileVersion") {
                    $fv = (Get-Item -Path $file.FullName).VersionInfo.FileVersion
                    if ($fv) { $versions += [string]$fv }
                }
            }
            if ($versions.Count -eq 0) {
                Write-Host ""; Write-Host "$($driver.Key), keine $fileName oder Version gefunden." -ForegroundColor Red; Write-Host ""
                continue
            }
            $highest = if ($versions.Count -eq 1) { $versions[0] } else { ($versions | Sort-Object { [Version]$_ } -Descending)[0] }
            $driverVersions += [PSCustomObject]@{ DriverName=$driver.Key; Versions=$highest; DirectoryPath=$cleanedDirs }
        }
        return $driverVersions
    }

    $driverVersions = GatherLocalDriverInfos
    Write-DeployLog -Message "Lokale Treiber-Infos: $($driverVersions.Count)" -Level 'INFO'

    # ── Compare online vs local ────────────────────────────────────────────────
    $driversToUpdate = @()
    foreach ($ld in $latestDrivers) {
        $local     = $driverVersions | Where-Object { $_.DriverName -eq $ld.DriverGroup }
        $localPath = $driverVersions | Where-Object { $_.DriverName -eq $ld.DriverGroup } | Select-Object -ExpandProperty DirectoryPath
        if ($local) {
            $onVer  = [Version]$ld.Version
            $lcVer  = [Version]$local.Versions
            Write-Host "########################################"
            Write-Host ""
            Write-Host "$($ld.DriverGroup)"
            Write-Host "    Lokale Version: $lcVer" -ForegroundColor Cyan
            Write-Host "    Online Version: $onVer" -ForegroundColor Cyan
            Write-DeployLog -Message "$($ld.DriverGroup): Lokal=$lcVer | Online=$onVer" -Level 'DEBUG'
            if ($onVer -gt $lcVer) {
                $driversToUpdate += [PSCustomObject]@{ DriverName=$ld.DriverGroup; Version=$ld.Version; DownloadLink=$ld.DownloadLink; DirectoryPath=$localPath }
                Write-Host "        $($ld.DriverGroup) Update gefunden!" -ForegroundColor Green
                Write-DeployLog -Message "Update: $($ld.DriverGroup) $lcVer → $onVer" -Level 'INFO'
            } else {
                Write-Host "        Kein Online Update verfügbar. $($ld.DriverGroup) ist aktuell." -ForegroundColor DarkGray
            }
            Write-Host ""
        }
    }
    Write-Host "########################################"; Write-Host ""
    Write-DeployLog -Message "Treiber zum Aktualisieren: $($driversToUpdate.Count)" -Level 'INFO'

    # ── Download + extract updates ─────────────────────────────────────────────
    function ExtractExeFile {
        $sz = "C:\Program Files\7-Zip\7z.exe"
        if (-not (Test-Path $sz)) { return }
        $base   = [System.IO.Path]::GetFileNameWithoutExtension($downloadFileName)
        $target = Join-Path $InstallationFolder $base
        if (-not (Test-Path $target)) { New-Item -Path $target -ItemType Directory | Out-Null }
        Write-Host "Die Datei wird extrahiert." -ForegroundColor Yellow
        & "$sz" x "$downloadPath" -o"$target" -y *> $null
    }

    if ($driversToUpdate.Count -gt 0) {
        $dlTemp = Join-Path ([System.IO.Path]::GetTempPath()) "DriverDownloads"
        if (-not (Test-Path $dlTemp)) { New-Item -Path $dlTemp -ItemType Directory | Out-Null }

        foreach ($drv in $driversToUpdate) {
            Write-Host ""; Write-Host "$($drv.DriverName) wird heruntergeladen.." -ForegroundColor Green; Write-Host ""
            $downloadUrl      = $drv.DownloadLink
            $downloadFileName = ([System.IO.Path]::GetFileName($downloadUrl)) -replace '\?.*$',''
            $downloadPath     = Join-Path $dlTemp $downloadFileName
            Write-DeployLog -Message "Download $($drv.DriverName): $downloadUrl" -Level 'INFO'

            Import-Module BitsTransfer -ErrorAction SilentlyContinue
            if (Get-Module -Name BitsTransfer -ListAvailable) {
                Start-BitsTransfer -Source $downloadUrl -Destination $downloadPath
            } else {
                Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath | Out-Null
            }

            if (Test-Path $downloadPath) {
                ExtractExeFile
                Remove-Item -Path $downloadPath -Force -ErrorAction SilentlyContinue

                $base         = [System.IO.Path]::GetFileNameWithoutExtension($downloadFileName)
                $targetFolder = Join-Path $InstallationFolder $base

                if ($drv.DriverName -eq "Armoury Crate") {
                    $acInstallPath  = "$InstallationFolder\ArmouryCrateInstallTool"
                    $acSub          = Get-ChildItem -Path $acInstallPath -Directory -Filter "ArmouryCrateInstaller_*" -ErrorAction SilentlyContinue
                    if ($acSub) {
                        $targetFolder = Join-Path $InstallationFolder $acSub.Name
                        Move-Item -Path $acSub.FullName -Destination $targetFolder
                        Remove-Item -Path $acInstallPath -Recurse -Force -ErrorAction SilentlyContinue
                        $acExe = Join-Path $targetFolder "ArmouryCrateInstaller.exe"
                        if (Test-Path $acExe) {
                            $svcVer = ($latestDrivers | Where-Object DriverName -eq 'Armoury Crate & Aura Creator Installer' | Select-Object -ExpandProperty ServiceVersion)
                            if ($svcVer) {
                                @{ ServiceVersion=$svcVer; Tool='ArmouryCrateInstaller'; StampedOn=(Get-Date).ToString('u') } |
                                    ConvertTo-Json | Set-Content "$acExe.json"
                            }
                        }
                    }
                }

                if (Test-Path $targetFolder) {
                    Remove-Item -Path $drv.DirectoryPath -Force -Recurse -ErrorAction SilentlyContinue
                    Write-Host ""; Write-Host "$($drv.DriverName) wurde aktualisiert.." -ForegroundColor Green
                    Write-DeployLog -Message "$($drv.DriverName) aktualisiert: $targetFolder" -Level 'SUCCESS'
                } else {
                    Write-Host "Download ist fehlgeschlagen. $($drv.DriverName) wurde nicht aktualisiert." -ForegroundColor Red
                    Write-DeployLog -Message "Zielordner fehlt: $targetFolder" -Level 'ERROR'
                }
            } else {
                Write-Host "Download ist fehlgeschlagen. $($drv.DriverName) wurde nicht aktualisiert." -ForegroundColor Red
                Write-DeployLog -Message "Download fehlgeschlagen: $downloadUrl" -Level 'ERROR'
            }
        }
        Remove-Item -Path $dlTemp -Recurse -Force -ErrorAction SilentlyContinue
    }

    # ── Installed driver check (WMI) ───────────────────────────────────────────
    Write-DeployLog -Message "Prüfe installierte Treiber via CIM." -Level 'INFO'
    $driverNames = @("Realtek Bluetooth Adapter","Realtek High Definition Audio","Realtek Gaming 2.5GbE Family Controller","Realtek 8852BE Wireless LAN WiFi 6 PCI-E NIC","AMD Radeon(TM) Graphics")
    $chipsetPaths = @(
        @{ Path="HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\AMD_Chipset_IODrivers"; ValueName="DisplayVersion" },
        @{ Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AMD_Chipset_IODrivers"; ValueName="DisplayVersion" }
    )
    $InstalledDrivers = @()
    foreach ($dn in $driverNames) {
        $inst = Get-CimInstance -ClassName Win32_PnPSignedDriver | Where-Object { $_.DeviceName -eq $dn }
        $InstalledDrivers += [PSCustomObject]@{ DriverName=$dn; InstalledVersion=if ($inst) { $inst.DriverVersion } else { "Not Installed" } }
    }
    $chipVer = $null
    foreach ($reg in $chipsetPaths) {
        if (Test-Path $reg.Path) {
            $v = Get-ItemProperty -Path $reg.Path -Name $reg.ValueName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $reg.ValueName -ErrorAction SilentlyContinue
            if ($v) { $chipVer = $v; break }
        }
    }
    $InstalledDrivers += [PSCustomObject]@{ DriverName="AMD Chipset Driver Suite"; InstalledVersion=if ($chipVer) { $chipVer } else { "Not Installed" } }

    $driverVersions = GatherLocalDriverInfos
    $nameMap = @{
        "Realtek Bluetooth Adapter"                    = "Bluetooth Driver"
        "Realtek High Definition Audio"                = "Audio Driver"
        "Realtek Gaming 2.5GbE Family Controller"      = "LAN Driver"
        "Realtek 8852BE Wireless LAN WiFi 6 PCI-E NIC" = "Wi-Fi Driver"
        "AMD Radeon(TM) Graphics"                      = "Graphics Driver"
        "AMD Chipset Driver Suite"                     = "Chipset Driver"
    }
    $newDriversToUpdate = @()
    foreach ($inst in $InstalledDrivers) {
        $mapped = $nameMap[$inst.DriverName]
        if (-not $mapped) { continue }
        $local = $driverVersions | Where-Object { $_.DriverName -eq $mapped }
        if (-not $local) { continue }
        Write-Host "$($local.DriverName) ist installiert." -ForegroundColor Green
        Write-Host "    Installierte Version:       $($inst.InstalledVersion)" -ForegroundColor Cyan
        Write-Host "    Installationsdatei Version: $($local.Versions)" -ForegroundColor Cyan
        if ($inst.InstalledVersion -ne "Not Installed") {
            if ([Version]$inst.InstalledVersion -lt [Version]$local.Versions) {
                $newDriversToUpdate += [PSCustomObject]@{ DriverName=$local.DriverName; InstalledVersion=$inst.InstalledVersion; DownloadedVersion=$local.Versions; DirectoryPath=$local.DirectoryPath }
                Write-Host "        Veraltete $($local.DriverName) ist installiert. Update wird gestartet." -ForegroundColor Magenta
                Write-DeployLog -Message "Treiber-Update nötig: $($local.DriverName)" -Level 'INFO'
            } else {
                Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
            }
        } else {
            Write-Host "        Treiber ist nicht installiert." -ForegroundColor Yellow
        }
        Write-Host ""
    }

    # ── Launch install scripts ─────────────────────────────────────────────────
    $DriversArgsString = ($newDriversToUpdate | ForEach-Object { "$($_.DriverName)|$($_.InstalledVersion)|$($_.DownloadedVersion)|$($_.DirectoryPath)" }) -join ','
    $ps1 = "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"

    if ($InstallationFlag) {
        Start-Process -FilePath $ps1 -ArgumentList @("-ExecutionPolicy Bypass","-NoLogo","-NoProfile","-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\TreiberAmdPcMainboardInstall.ps1`"","-DriversToUpdate `"$DriversArgsString`"","-InstallationFlag") -NoNewWindow -Wait
        Start-Process -FilePath $ps1 -ArgumentList @("-ExecutionPolicy Bypass","-NoLogo","-NoProfile","-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\ArmouryCrateInstall.ps1`"","-InstallationFlag") -NoNewWindow -Wait
    } elseif ($newDriversToUpdate.Count -gt 0) {
        Start-Process -FilePath $ps1 -ArgumentList @("-ExecutionPolicy Bypass","-NoLogo","-NoProfile","-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\TreiberAmdPcMainboardInstall.ps1`"","-DriversToUpdate `"$DriversArgsString`"") -NoNewWindow -Wait
    }
    if (-not $InstallationFlag) {
        Start-Process -FilePath $ps1 -ArgumentList @("-ExecutionPolicy Bypass","-NoLogo","-NoProfile","-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\ArmouryCrateInstall.ps1`"") -NoNewWindow -Wait
    }

    Write-Host ""
    Write-DeployLog -Message "Treiber-Workflow abgeschlossen." -Level 'SUCCESS'

} else {
    Write-Host ""
    Write-Host "        Treiber sind NICHT für dieses System geeignet." -ForegroundColor Blue
    Write-Host ""
    Write-DeployLog -Message "System $PCName ist nicht Zielsystem. Keine Aktionen." -Level 'INFO'
}

Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
