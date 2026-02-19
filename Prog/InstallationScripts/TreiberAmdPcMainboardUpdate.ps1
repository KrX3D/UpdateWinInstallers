# Usage: .\TreiberAmdPcMainboardUpdate.ps1 -Model "PRIME X670-P WiFi" -OS "Windows 10 64-bit" -ShowBrowser $false
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

Write_LogEntry -Message "Script gestartet mit InstallationFlag: $($InstallationFlag), Model: $($Model), OS: $($OS), ShowBrowser: $($ShowBrowser)" -Level "INFO"
Write_LogEntry -Message "ProgramName: $($ProgramName); ScriptType: $($ScriptType)" -Level "DEBUG"

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
Write_LogEntry -Message "Konfigurationspfad gesetzt: $($configPath)" -Level "DEBUG"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
    Write_LogEntry -Message "Konfigurationsdatei gefunden und importiert: $($configPath)" -Level "INFO"
} else {
    Write_LogEntry -Message "Konfigurationsdatei nicht gefunden: $($configPath)" -Level "ERROR"
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    pause
    exit
}

$InstallationFolder = "$NetworkShareDaten\Treiber\AMD_PC"
Write_LogEntry -Message "Installations-Ordner: $($InstallationFolder)" -Level "DEBUG"

#Check Pc, Skipp if not AMD PC
$scriptPath = "$Serverip\Daten\Prog\InstallationScripts\GetPcByUUID.ps1"
Write_LogEntry -Message "Script zur PC-Ermittlung: $($scriptPath)" -Level "DEBUG"

try {
	#$PCName = & $scriptPath
	$PCName = & "powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Verbose:$false
    Write_LogEntry -Message "PC-Ermittlungs-Script erfolgreich ausgeführt. PCName: $($PCName)" -Level "SUCCESS"
} catch {
    Write_LogEntry -Message "Fehler beim Laden des Scripts $($scriptPath): $($_)" -Level "ERROR"
    Write-Host "Failed to load script $scriptPath. Reason: $($_.Exception.Message)" -ForegroundColor Red
    Pause
    Exit
}

# Print the pc name
Write-Host "PC Name: $PCName"
Write_LogEntry -Message "PC Name: $($PCName) ausgegeben." -Level "DEBUG"

if($PCName -eq "KrX-AMD-PC")
{
	############################################################################
	#Gather Driver Information Start
	############################################################################
	Write_LogEntry -Message "Zielsystem erkannt: $($PCName). Beginne Treiber-Informationssammlung." -Level "INFO"
	
	# Create temp file for JavaScript code
	$TempFolder = [System.IO.Path]::GetTempPath()
	$JsFile = Join-Path -Path $TempFolder -ChildPath "asus_driver_downloader.js"
	Write_LogEntry -Message "TempFolder: $($TempFolder); JsFile: $($JsFile)" -Level "DEBUG"

	# Check if Node.js is installed
	if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
		Write_LogEntry -Message "Node.js nicht installiert. Abbruch." -Level "ERROR"
		Write-Host "Node.js is not installed. Installing Node.js." -ForegroundColor Yellow
		& $PSHostPath -ExecutionPolicy Bypass -NoLogo -NoProfile -File "$Serverip\Daten\Prog\InstallationScripts\NodeUpdate.ps1" -InstallationFlag
		
		$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
		#Write-Host "Node.js was freshly installed and cant be used, updating next time. Exiting." -ForegroundColor Yellow
		#exit 1
	} else {
		Write_LogEntry -Message "Node.js ist installiert." -Level "INFO"
	}
	
	# ——————————————————————————————————————————
	# Ensure Puppeteer is installed & up-to-date
	# ——————————————————————————————————————————

	# Path where your asus_download.js (and package.json) live.
	# If you don’t have a package.json there yet, you can create one via `npm init -y`
	$npmWorkingFolder = $env:USERPROFILE
	Write_LogEntry -Message "NPM Working Folder: $($npmWorkingFolder)" -Level "DEBUG"

	Push-Location $npmWorkingFolder

	# 1) Do we have puppeteer?
	$puppeteerInstalled = $false
	try {
		# npm list will throw if not installed
		npm list puppeteer --depth=0 | Out-Null
		$puppeteerInstalled = $true
		Write_LogEntry -Message "Puppeteer ist installiert (lokal)." -Level "DEBUG"
	} catch {
		$puppeteerInstalled = $false
		Write_LogEntry -Message "Puppeteer nicht installiert." -Level "DEBUG"
	}

	if (-not $puppeteerInstalled) {
		Write_LogEntry -Message "Puppeteer nicht gefunden - Installation gestartet." -Level "INFO"
		Write-Host "Puppeteer not found-installing..." -ForegroundColor Yellow
		npm install puppeteer
		Write_LogEntry -Message "Puppeteer Installation abgeschlossen." -Level "SUCCESS"
	} else {
		# 2) Check whether it’s outdated
		$outdated = npm outdated puppeteer --parseable
		if ($outdated) {
			Write_LogEntry -Message "Puppeteer veraltet - Update gestartet." -Level "INFO"
			Write-Host "Puppeteer is out-of-date-updating..." -ForegroundColor Yellow
			npm update puppeteer
			Write_LogEntry -Message "Puppeteer Update abgeschlossen." -Level "SUCCESS"
		} else {
			Write_LogEntry -Message "Puppeteer ist aktuell." -Level "DEBUG"
			Write-Host "Puppeteer is already up-to-date." -ForegroundColor Green
		}
	}

	Pop-Location

# JavaScript code with parameter support
$JsCode = @'
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

// Get command line arguments
const args = process.argv.slice(2);
const model = args[0] || "PRIME X670-P WiFi";
const os = args[1] || "Windows 10 64-bit";
const showBrowser = args[2] === "true";

// Map OS names to ASUS OS IDs
const osIdMap = {
  "Windows 10 64-bit": "45",
  "Windows 11 64-bit": "52"
};

/**
 * Extract ASUS driver downloads for a specific model
 * @param {string} model - The ASUS model (e.g., "PRIME X670-P WiFi")
 * @param {string} os - Operating system "Windows 10 64-bit" or "Windows 11 64-bit"
 */
async function extractASUSDrivers(model, os) {
  const browser = await puppeteer.launch({
    headless: !showBrowser, // Run headless by default, show browser if specified
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    defaultViewport: { width: 1920, height: 1080 }
  });
  
  try {
    // console.log(`Starting extraction for ${model} - ${os}...`);
    const page = await browser.newPage();
    
    // Get the correct OS ID for the selected OS
    const osId = osIdMap[os] || "45"; // Default to Windows 10 if not found
    // console.log(`Using OS ID: ${osId} for ${os}`);
    
    // Format model name for URL
    const formattedModel = model.replace(/ /g, '_');
    const modelLower = formattedModel.toLowerCase();
    
    // Go directly to the API URL with the correct OS ID
    const apiUrl = `https://www.asus.com/support/api/product.asmx/GetPDDrivers?website=de&model=${modelLower}&pdhashedid=&cpu=&osid=${osId}&pdid=99999&siteID=www&sitelang=`;
    // console.log(`Directly accessing API: ${apiUrl}`);
    
    // Access the API directly
    const apiResponse = await page.goto(apiUrl, { waitUntil: 'networkidle2', timeout: 60000 });
    // console.log("API request completed");
    
    // Get the API response as text
    const apiText = await apiResponse.text();
    // console.log(`API response length: ${apiText.length}`);
    
    // Save the raw API response for debugging
    const osForFilename = os.replace(/\s+/g, '_');
    fs.writeFileSync(
      path.join(require('os').tmpdir(), `asus_api_raw_${osForFilename}.txt`),
      apiText
    );
    
    // Parse the API response
    try {
      const apiJson = JSON.parse(apiText);
      // console.log("Successfully parsed API response");
      
      // Save the parsed API response
      fs.writeFileSync(
        path.join(require('os').tmpdir(), `asus_api_${osForFilename}.json`),
        JSON.stringify(apiJson, null, 2)
      );
      
      // Process the API response to extract driver info
      const drivers = processApiResponse(apiJson);
      
      // Save the formatted driver list
      const outputPath = path.join(require('os').tmpdir(), `ASUS_${model.replace(/\s+/g, '_')}_${osForFilename}_drivers.json`);
      fs.writeFileSync(outputPath, JSON.stringify(drivers, null, 2));
      // console.log(`Extracted ${drivers.length} drivers. Saved to: ${outputPath}`);
      
      // If no drivers were found, get the web page content for fallback
      if (drivers.length === 0) {
        // console.log("No drivers found in API response. Trying webpage extraction as fallback...");
        await fallbackWebpageExtraction(page, model, os, formattedModel);
      }
      
      return drivers;
    } catch (e) {
      // console.log("Failed to parse API response:", e.message);
      // console.log("Falling back to webpage extraction...");
      await fallbackWebpageExtraction(page, model, os, formattedModel);
      return [];
    }
  } catch (error) {
    console.error("Script failed with error:", error);
    fs.writeFileSync(
      path.join(require('os').tmpdir(), 'asus_error.txt'),
      String(error.stack || error),
      'utf8'
    );
    return [];
  } finally {
    await browser.close();
    // console.log("Browser closed.");
  }
}

/**
 * Fallback method to extract drivers from the webpage
 */
async function fallbackWebpageExtraction(page, model, os, formattedModel) {
  try {
    const url = `https://www.asus.com/de/supportonly/${formattedModel}/helpdesk_download`;
    // console.log(`Navigating to ${url}...`);
    
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });
    // console.log("Page loaded.");
    
    // Select the OS (Windows 10 or Windows 11)
    try {
      // console.log("Looking for OS selector...");
      await page.waitForSelector('.SortSelect__selectContent__1u-8d', { timeout: 15000 });
      
      // Click on the OS dropdown
      await page.click('.SortSelect__selectContent__1u-8d');
      // console.log("Clicked OS dropdown");
      
      // Wait for options to appear
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // Get the current selected OS
      const currentOS = await page.evaluate(() => {
        const osElement = document.querySelector('.SortSelect__selectContent__1u-8d');
        return osElement ? osElement.textContent.trim() : null;
      });
      
      // console.log(`Current selected OS: ${currentOS}`);
      
      // Only change the OS if it's not already what we want
      if (currentOS !== os) {
        // Select Windows version based on parameter (using exact OS string from the website)
        // console.log(`Looking for OS option: ${os}`);
        
        // First try exact match
        let winSelector = `div[aria-label="${os}"]`;
        let winOption = await page.$(winSelector);
        
        // If exact match fails, try contains match
        if (!winOption) {
          winSelector = `div[aria-label*="${os}"]`;
          winOption = await page.$(winSelector);
        }
        
        if (winOption) {
          await winOption.click();
          // console.log(`Selected ${os}`);
          
          // Wait for content to refresh
          await new Promise(resolve => setTimeout(resolve, 5000));
        } else {
          // console.log(`${os} option not found, staying with default selection`);
        }
      } else {
        // console.log(`${os} is already selected, no need to change`);
      }
    } catch (err) {
      // console.log("OS selection failed:", err.message);
      // console.log("Continuing with default OS selection...");
    }
    
    // Take a screenshot for verification
    const osForFilename = os.replace(/\s+/g, '_');
    const screenshotPath = path.join(require('os').tmpdir(), `asus_page_${osForFilename}.png`);
    await page.screenshot({ path: screenshotPath, fullPage: true });
    // console.log(`Screenshot saved: ${screenshotPath}`);
    
    // Scroll to load all content and any lazy-loaded API calls
    await autoScroll(page);
    
    // Wait a bit more to ensure all content is loaded
    await new Promise(resolve => setTimeout(resolve, 5000));
    
    // Extract drivers from page content
    const pageDrivers = await extractDriversFromPage(page);
    
    if (pageDrivers.length > 0) {
      const fallbackPath = path.join(require('os').tmpdir(), `ASUS_${model.replace(/\s+/g, '_')}_${osForFilename}_drivers_fallback.json`);
      fs.writeFileSync(fallbackPath, JSON.stringify(pageDrivers, null, 2));
      // console.log(`Extracted ${pageDrivers.length} drivers via fallback method. Saved to: ${fallbackPath}`);
    } else {
      // console.log("No drivers found on page.");
    }
    
    // Save full HTML for offline analysis if needed
    const htmlPath = path.join(require('os').tmpdir(), `asus_page_${osForFilename}_full.html`);
    const html = await page.content();
    fs.writeFileSync(htmlPath, html);
    // console.log(`Full HTML saved to: ${htmlPath}`);
  } catch (err) {
    console.error("Fallback webpage extraction failed:", err.message);
  }
}

/**
 * Process a single API response to extract driver information
 */
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
      // try each regex until one matches
      const desc = file.Description || "";
      let svcVer = null;
      for (const re of vReList) {
        const m = desc.match(re);
        if (m && m[1]) {
          svcVer = m[1];
          break;
        }
      }

      drivers.push({
        category:       category.Name                 || "Unknown",
        name:           file.Title                    || "Unknown",
        version:        file.Version                  || "Unknown",
        serviceVersion: svcVer,
        fileSize:       file.FileSize                 || "Unknown",
        releaseDate:    file.ReleaseDate              || "Unknown",
        downloadUrl:    file.DownloadUrl?.Global      || "Unknown",
        id:             file.Id                       || "Unknown"
      });
    });
  });

  return drivers;
}

/**
 * Fallback method to extract drivers directly from page content
 */
async function extractDriversFromPage(page) {
  try {
    return await page.evaluate(() => {
      const drivers = [];
      
      // Find all driver boxes
      const driverBoxes = document.querySelectorAll('.productSupportDriverBIOSBox, .ProductSupportDriverBIOS__productSupportDriverBIOSBox__ihsCB');
      
      driverBoxes.forEach(box => {
        // Extract title
        const titleEl = box.querySelector('.ProductSupportDriverBIOS__fileTitle__GE44d');
        const title = titleEl ? titleEl.textContent.trim() : "Unknown";
        
        // Extract info container
        const infoContainer = box.querySelector('.ProductSupportDriverBIOS__fileInfo__2c5GN');
        
        // Extract version
        let version = "Unknown";
        const versionEl = infoContainer ? infoContainer.querySelector('div:first-child') : null;
        if (versionEl && versionEl.textContent.includes('Version')) {
          version = versionEl.textContent.replace('Version', '').trim();
        }
        
        // Extract file size
        let fileSize = "Unknown";
        const sizeEl = infoContainer ? infoContainer.querySelector('.ProductSupportDriverBIOS__fileSize__eticu') : null;
        if (sizeEl) {
          fileSize = sizeEl.textContent.trim();
        }
        
        // Extract release date
        let releaseDate = "Unknown";
        const dateEl = infoContainer ? infoContainer.querySelector('.ProductSupportDriverBIOS__releaseDate__3o309') : null;
        if (dateEl) {
          releaseDate = dateEl.textContent.trim();
        }
        
        // Try to find category
        let category = "Unknown";
        const parentSection = box.closest('[class*="type-"]');
        if (parentSection) {
          const categoryEl = parentSection.querySelector('h2, .title');
          if (categoryEl) {
            category = categoryEl.textContent.trim();
          }
        }
        
        // Try to find download link from button
        let downloadUrl = "Unknown";
        const downloadBtn = box.querySelector('.ProductSupportDriverBIOS__downloadBtn__204JI, .SolidButton__btn__1NmTw');
        if (downloadBtn) {
          downloadUrl = downloadBtn.getAttribute('href') || 
                       downloadBtn.getAttribute('data-href') ||
                       downloadBtn.getAttribute('data-url') ||
                       downloadBtn.getAttribute('data-downloadurl') ||
                       "Unknown";
        }
        
        drivers.push({
          category,
          name: title,
          version,
          fileSize,
          releaseDate,
          downloadUrl,
          extractionMethod: "page_content"
        });
      });
      
      return drivers;
    });
  } catch (err) {
    console.error("Error extracting drivers from page:", err.message);
    return [];
  }
}

/**
 * Helper function to scroll page to load all content
 */
async function autoScroll(page) {
  await page.evaluate(async () => {
    await new Promise((resolve) => {
      let totalHeight = 0;
      const distance = 100;
      const timer = setInterval(() => {
        const scrollHeight = document.body.scrollHeight;
        window.scrollBy(0, distance);
        totalHeight += distance;
        
        if (totalHeight >= scrollHeight) {
          clearInterval(timer);
          resolve();
        }
      }, 100);
    });
  });
}

// Run the extraction with parameters passed from PowerShell
extractASUSDrivers(model, os).then(() => {
  // console.log("Extraction complete!");
});
'@

	# Write JS file to temp folder
	Set-Content -Path $JsFile -Value $JsCode -Encoding UTF8
	Write_LogEntry -Message "JavaScript Datei erstellt: $($JsFile)" -Level "DEBUG"
	#Write-Host "JavaScript file created: $JsFile"

	# Run the JS file with Node.js, passing the parameters
	#Write-Host "Executing Puppeteer script to fetch driver information..." -ForegroundColor Cyan
	Write-Host "Model: $Model" -ForegroundColor Yellow
	Write-Host "OS: $OS" -ForegroundColor Yellow
	if ($ShowBrowser) {
		Write-Host "NOTE: A browser window will open - do not close it until the script finishes." -ForegroundColor Yellow
		Write_LogEntry -Message "Puppeteer wird sichtbar ausgeführt (ShowBrowser=$($ShowBrowser))." -Level "INFO"
	} else {
		Write-Host "Running in headless mode - no browser window will be shown." -ForegroundColor Yellow
		Write_LogEntry -Message "Puppeteer wird headless ausgeführt (ShowBrowser=$($ShowBrowser))." -Level "INFO"
	}
	Write-Host ""

	# Execute node with parameters
	Write_LogEntry -Message "Starte Node.js Script: $($JsFile) mit Parametern Model=$($Model), OS=$($OS), ShowBrowser=$($ShowBrowser)" -Level "INFO"
	node $JsFile "$Model" "$OS" "$ShowBrowser"
	Write_LogEntry -Message "Node.js Script beendet: $($JsFile)" -Level "INFO"

	# After script completes, find and display the results file
	$OSForFilename = $OS.Replace(" ", "_")
	$ResultFile = Get-ChildItem -Path $TempFolder -Filter "ASUS_$($Model.Replace(' ', '_'))_${OSForFilename}_drivers.json" | Select-Object -First 1
	if ($ResultFile) {
		Write_LogEntry -Message "Driver information saved to: $($ResultFile.FullName)" -Level "SUCCESS"
		#Write-Host "Driver information saved to: $($ResultFile.FullName)" -ForegroundColor Green
		
		# Optionally display driver count
		$DriversJson = Get-Content $ResultFile.FullName -Raw | ConvertFrom-Json
		Write-Host "Found $($DriversJson.Count) drivers for $Model ($OS)" -ForegroundColor Green
		Write_LogEntry -Message "Gefundene Treiberanzahl: $($DriversJson.Count) für Model $($Model) OS $($OS)" -Level "INFO"
	} else {
		# Check for fallback result
		$FallbackFile = Get-ChildItem -Path $TempFolder -Filter "ASUS_$($Model.Replace(' ', '_'))_${OSForFilename}_drivers_fallback.json" | Select-Object -First 1
		if ($FallbackFile) {
			Write_LogEntry -Message "Driver information saved to (fallback method): $($FallbackFile.FullName)" -Level "WARNING"
			Write-Host "Driver information saved to (fallback method): $($FallbackFile.FullName)" -ForegroundColor Yellow
			
			# Optionally display driver count
			$DriversJson = Get-Content $FallbackFile.FullName -Raw | ConvertFrom-Json
			Write-Host "Found $($DriversJson.Count) drivers for $Model ($OS)" -ForegroundColor Yellow
			Write_LogEntry -Message "Gefundene Treiberanzahl (Fallback): $($DriversJson.Count) für Model $($Model) OS $($OS)" -Level "INFO"
		} else {
			Write_LogEntry -Message "Keine Treiber-Informationsdatei gefunden. Prüfe Fehlerdatei." -Level "ERROR"
			Write-Host "No driver information file found. Check for errors." -ForegroundColor Red
			
			# Check for error file
			$ErrorFile = Join-Path -Path $TempFolder -ChildPath "asus_error.txt"
			if (Test-Path $ErrorFile) {
				Write-Host "Error details:" -ForegroundColor Red
				Get-Content $ErrorFile | Write-Host -ForegroundColor Red
				Write_LogEntry -Message "Fehlerdatei vorhanden: $($ErrorFile)" -Level "ERROR"
			}
		}
	}

	# Remove JS file after execution
	Remove-Item -Path $JsFile -Force
	Write_LogEntry -Message "JavaScript Datei gelöscht: $($JsFile)" -Level "DEBUG"
	#Write-Host "JavaScript file deleted."
	Remove-Item -Path (Get-ChildItem -Path $TempFolder -Filter "asus_api_*.*").FullName -Force
	Write_LogEntry -Message "Temporäre API Dateien entfernt." -Level "DEBUG"

	# Read the downloaded JSON content
	$outputPath = (Get-ChildItem -Path $TempFolder -Filter "ASUS_*_drivers.json").FullName
	if (Test-Path $outputPath) {
	  $jsonContent = Get-Content -Path $outputPath -Raw
	  $allDrivers    = $jsonContent | ConvertFrom-Json
	  Write_LogEntry -Message "ASUS JSON geladen: $($outputPath); Einträge: $($allDrivers.Count)" -Level "INFO"
	} else {
	  Write_LogEntry -Message "ASUS download JSON nicht gefunden: $($outputPath)" -Level "ERROR"
	  Write-Warning "ASUS download JSON not found at $outputPath; skipping parsing."
	  return
	}
	
	# Remove the specific file if it exists
	if (Test-Path -Path $outputPath) {
		Remove-Item -Path $outputPath -Force
		Write_LogEntry -Message "ASUS JSON temporäre Datei entfernt: $($outputPath)" -Level "DEBUG"
	}

	$excludeKeywords = @(
	  "AI Suite","WinRAR","MY ASUS","CPU-Z",
	  "Armoury Crate Uninstall Tool","for Ryzen",
	  "RAIDXpert2","Aura Creator Installer"
	)
	Write_LogEntry -Message "Exclude-Keywords gesetzt: $($excludeKeywords -join ', ')" -Level "DEBUG"

	# filtere alle Einträge, deren name keines der Keywords enthält	
	$filtered = $allDrivers | Where-Object {
		$item = $_
		$exclude = $false
		
		foreach ($keyword in $excludeKeywords) {
			if ($item.name -like "*$keyword*") {
				$exclude = $true
				break
			}
		}
		
		# Special case: Keep the Armoury Crate Installer
		if ($item.name -eq "Armoury Crate & Aura Creator Installer") {
			$exclude = $false
		}
		
		-not $exclude
	}
	Write_LogEntry -Message "Gefilterte Treiberanzahl: $($filtered.Count)" -Level "DEBUG"

	function Get-BaseName {
	  param($name)
	  # alles ab dem Versions-Pattern (z.B. " V6.10..." oder " v7.01...") sowie OS-Angaben abschneiden
	  if ($name -match '^(.*?)\s+[Vv]\d') {
		return $Matches[1].Trim()
	  }
	  # Falls kein " V<Zahl>" gefunden, alles vor " for Windows" nehmen
	  if ($name -match '^(.*?)\s+for\s+Windows') {
		return $Matches[1].Trim()
	  }
	  return $name.Trim()
	}
	
	$latestPerGroup = $filtered |
	  # füge ein ScriptProperty BaseName hinzu
	  Select-Object *, @{
		Name='BaseName'; 
		Expression={ Get-BaseName $_.name }
	  } |
	  # gruppiere nach BaseName
	  Group-Object -Property BaseName |
	  ForEach-Object {
		# Innerhalb jeder Gruppe nach releaseDate (Datum) und version sortieren
		$_.Group |
		  Sort-Object @{ Expression = { [datetime]($_.releaseDate) } }, `
					   @{ Expression = { [version]($_.version) } } -Descending |
		  Select-Object -First 1
	  }

	Write_LogEntry -Message "Anzahl latestPerGroup: $($latestPerGroup.Count)" -Level "DEBUG"

	#$latestPerGroup | Select-Object category, name, version, serviceVersion, fileSize, releaseDate, downloadUrl | ConvertTo-Json -Depth 4
	
	# Initialize an empty array to hold the extracted data
	$downloadData = @()
	
	$downloadData = $latestPerGroup | ForEach-Object {
		[PSCustomObject]@{
			Label        = $_.name
			DriverName   = $_.BaseName
			Version      = $_.version
			ServiceVersion = $_.serviceVersion
			DownloadLink = $_.downloadUrl
			releaseDate  = $_.releaseDate
		}
	}
	Write_LogEntry -Message "DownloadData aufgebaut: $($downloadData.Count) Einträge." -Level "DEBUG"

	# Define regex patterns for grouping similar drivers
	$driverPatterns = @{
		"Graphics Driver"  = "AMD\s*Graphics|VGA"           # AMD Graphics and VGA
		"RAID Driver"      = "RAID"              		   # RAID
		"Bluetooth Driver" = "Bluetooth"                   # Bluetooth drivers
		"LAN Driver"       = "LAN|Ethernet"                 # LAN drivers
		"Wi-Fi Driver"     = "Wi[-]?Fi|Wireless"            # Wi-Fi and Wi-Fi related drivers
		"Audio Driver"     = "Audio"                        # Audio drivers
		"Armoury Crate"    = "Armoury\s*Crate"              # Armoury Crate drivers
		"Chipset Driver"   = "AMD\s*Chipset"           # AMD Graphics and VGA
	}
	
	$downloadData | ForEach-Object {
		$grp = "Other"
		foreach($key in $driverPatterns.Keys) {
			if ($_.DriverName -match $driverPatterns[$key]) {
				$grp = $key
				break
			}
		}
		$_ | Add-Member -MemberType NoteProperty -Name DriverGroup -Value $grp -Force
	}

	$latestDrivers = $downloadData |
		Group-Object -Property DriverGroup |
		ForEach-Object {
			$_.Group |
				Sort-Object `
				  @{Expression = { $_.releaseDate }}, `
				  @{Expression = { $_.Version    }} `
				  -Descending |
				Select-Object -First 1
		}
		
	#$latestDrivers | Sort-Object -Property DriverGroup | Format-Table DriverGroup, DriverName, Version, ServiceVersion, releaseDate, DownloadLink -AutoSize

	Write_LogEntry -Message "LatestDrivers bestimmt: $($latestDrivers.Count) Gruppen." -Level "DEBUG"

	############################################################################
	#Gather Driver Information End
	############################################################################

	############################################################################
	#Gather Local Driver Information Start
	############################################################################

	Function GatherLocalDriverInfos(){
		$localdriverPatterns = @{
			"Graphics Driver" = @{
				Path      = "$InstallationFolder\DRV_VGA_AMD_*"
				File      = "AsusSetup.ini"
				Section   = "InstallInfo"
				Key       = "Version"
				Delimiter = "false"
			}
			"RAID Driver" = @{ #skip since not used
				Path      = "$InstallationFolder\DRV_RAID_AMD_RAID_Driver_*"
				File      = "*\*.inf"
				Section   = "Version"
				Key       = "DriverVer"
				Delimiter = ","
			}
			"Bluetooth Driver" = @{
				Path      = "$InstallationFolder\DRV_BT_RTK_8852BE_*"
				File      = "AsusSetup.ini"
				Section   = "InstallInfo"
				Key       = "Version"
				Delimiter = "false"
			}
			"LAN Driver" = @{
				#Path      = "$InstallationFolder\DRV_LAN_Realtek_8125_SZ-TSD_W11_64_*"
				Path      = "$InstallationFolder\DRV_LAN_*"
				#File      = "*\AsusSetup.ini" #Unterordner z.B. Win10 und Win11 wo dann die INI drin liegt
				File      = "AsusSetup.ini"
				Section   = "InstallInfo"
				Key       = "Version"
				Delimiter = "false"
			}
			"Wi-Fi Driver" = @{
				Path      = "$InstallationFolder\DRV_WiFi_RTK_8852BE_*"
				File      = "AsusSetup.ini"
				Section   = "InstallInfo"
				Key       = "Version"
				Delimiter = "false"
			}
			"Audio Driver" = @{
				Path      = "$InstallationFolder\DRV_Audio_RTK_*"
				File      = "AsusSetup.ini"
				Section   = "InstallInfo"
				Key       = "Version"
				Delimiter = "false"
			}
			"Armoury Crate" = @{
				Path      = "$InstallationFolder\ArmouryCrateInstaller_*"
				File      = "ArmouryCrateInstaller.exe"
				Section   = ""
				Key       = ""
				Delimiter = "FileVersion"
			}
			"Chipset Driver" = @{
				#Path      = "$InstallationFolder\DRV_Chipset_AMD_AM5_TP_TSD_W11_64_*"
				#Path      = "$InstallationFolder\DRV_Chipset_*\Chipset"
				Path      = "$InstallationFolder\DRV_Chipset_*"
				#File      = "AsusSetup.ini" #direkt im root Ordner
				File      = "AsusSetup.ini"
				Section   = "InstallInfo"
				Key       = "Version"
				Delimiter = "false"
			}
		}

		# Initialize an array to store driver name and version objects
		$driverVersions = @()

		foreach ($driver in $localdriverPatterns.GetEnumerator()) {
			$driverName = $driver.Key
			$driverConfig = $driver.Value
			$pathPattern = $driverConfig.Path
			$filePattern = $driverConfig.File
			$section = $driverConfig.Section
			$key = $driverConfig.Key
			$delimiter = $driverConfig.Delimiter

			# Combine Path and File pattern to form the full search pattern
			$fullPattern = [System.IO.Path]::Combine($pathPattern, $filePattern)

			# Extract base path and wildcard folder pattern
			$basePath = Split-Path -Path $fullPattern
			$wildcardFolder = (Split-Path -Leaf (Split-Path -Path $fullPattern -Parent))
			$fileName = Split-Path -Leaf $fullPattern

			# Find directories that match the wildcard folder
			$matchedDirectories = Get-ChildItem -Path $basePath -Directory -Filter $wildcardFolder -ErrorAction SilentlyContinue
			# Remove trailing \* if present
			$cleanBasePath = if ($basePath -match '\\\*$') { $basePath -replace '\\\*$', '' } else { $basePath }
			$cleanedDirectories = Get-ChildItem -Path $cleanBasePath -Directory -Filter $wildcardFolder -ErrorAction SilentlyContinue
			$cleanedDirectories = if ($cleanedDirectories -match '\\Chipset$') { $cleanedDirectories -replace '\\Chipset$', '' } else { $cleanedDirectories }

			$matchedFiles = @()
			foreach ($dir in $matchedDirectories) {
				# Look for the exact file in each matching directory
				$files = Get-ChildItem -Path $dir.FullName -Filter $fileName -ErrorAction SilentlyContinue
				$matchedFiles += $files
			}

			# If no matching files are found, continue to the next driver
			if ($null -eq $matchedFiles) {
				#Write-Host "No matching files found for driver: $driverName. File pattern: $fullPattern" -ForegroundColor Red
				continue
			}

			# Initialize an array to store version numbers for the current driver
			$versions = @()

			# Process matched files
			foreach ($file in $matchedFiles) {
				if ($file.Extension -eq ".ini" -or $file.Extension -eq ".inf") {
					$iniContent = Get-Content -Path $file.FullName
					$inSection = $false
					foreach ($line in $iniContent) {
						if ($line -match "^\[$section\]") {
							$inSection = $true
							continue
						}
						if ($inSection -and $line -match "^\[") {
							break
						}
						if ($inSection -and $line -match "^\s*$key\s*=\s*(.+)") {
							$value = $matches[1]

							# Check if delimiter is specified and process accordingly
							if ($delimiter -and $delimiter -ne "false") {
								$value = ($value -split $delimiter)[-1]
							}

							# Ensure we add the value as a string to the versions array
							$versions += [string]$value
						}
					}
				} elseif ($file.Extension -eq ".exe") {
					# Process EXE file to get FileVersion
					if ($delimiter -eq "FileVersion") {
						$fileVersion = (Get-Item -Path $file.FullName).VersionInfo.FileVersion
						if ($fileVersion) {
							$versions += [string]$fileVersion
						}
					}
				}
			}

			# If no versions were found, output a message
			if ($versions.Count -eq 0) {
				Write-Host ""
				Write-Host "$driverName, keine $fileName oder Version gefunden." -ForegroundColor Red
				Write-Host ""
				continue
			}

			# Only select the highest version from the versions found
			if ($versions.Count -gt 0) {
				$highestVersion = if ($versions.Count -eq 1) {
					$versions[0]  # If there's only one version, use it directly
				} else {
					($versions | Sort-Object { [Version]$_ } -Descending)[0]  # Otherwise, sort and pick the highest
				}

				# Store the driver name and its highest version in an object, then add it to the array
				$driverVersions += [PSCustomObject]@{
					DriverName = $driverName
					Versions   = $highestVersion
					DirectoryPath = $cleanedDirectories
				}
			}
		}

		return $driverVersions
		# Output the driver versions
		#$driverVersions | Sort-Object -Property DriverName | Format-Table -AutoSize
	}

	$driverVersions = GatherLocalDriverInfos #Call Function to get Local Driver Infos
	Write_LogEntry -Message "Lokale DriverInfos gesammelt: $($driverVersions.Count) Einträge." -Level "DEBUG"

	############################################################################
	#Gather Local Driver Information End
	############################################################################

	############################################################################
	# Compare Online Drivers and Versions with local one Start
	############################################################################
	Write_LogEntry -Message "Vergleiche Online-Treiber mit lokalen Treibern." -Level "INFO"

	# Initialize an array to store drivers with newer versions
	$driversToUpdate = @()

	# Compare drivers between $latestDrivers and $driverVersions
	foreach ($latestDriver in $latestDrivers) {
		# Find the corresponding local driver by matching DriverName
		$localDriver = $driverVersions | Where-Object { $_.DriverName -eq $latestDriver.DriverGroup }
		$localDriverPath = $driverVersions | Where-Object { $_.DriverName -eq $latestDriver.DriverGroup } | Select-Object -ExpandProperty DirectoryPath
		
		if ($localDriver) {
			# If local driver exists, compare versions
			$latestVersion = [Version]$latestDriver.Version
			$localVersion = [Version]$localDriver.Versions
			
			# Output local and online versions
			Write-Host "########################################"
			Write-Host ""
			Write-Host "$($latestDriver.DriverGroup)"
			Write-Host "	Lokale Version: $localVersion" -ForegroundColor "Cyan"
			Write-Host "	Online Version: $latestVersion" -ForegroundColor "Cyan"
			Write_LogEntry -Message "Vergleich $($latestDriver.DriverGroup): Lokal=$($localVersion); Online=$($latestVersion)" -Level "DEBUG"

			# Check if the latest version is newer than the local version
			if ($latestVersion -gt $localVersion) {			
				# Add to the array if the latest version is newer
				$driversToUpdate += [PSCustomObject]@{
					DriverName   = $latestDriver.DriverGroup
					Version      = $latestDriver.Version
					DownloadLink = $latestDriver.DownloadLink
					DirectoryPath = $localDriverPath
				}

				# Output the update decision
				Write-Host "		$($latestDriver.DriverGroup) Update gefunden!" -ForegroundColor "Green"
				Write_LogEntry -Message "Update gefunden für $($latestDriver.DriverGroup): Online-Version $($latestVersion) neuer als lokal $($localVersion)" -Level "INFO"
			} else {
				Write-Host "		Kein Online Update verfügbar. $($latestDriver.DriverGroup) is aktuell." -ForegroundColor "DarkGray"
				Write_LogEntry -Message "Kein Update für $($latestDriver.DriverGroup)." -Level "DEBUG"
			}
			Write-Host ""
		}
	}
	Write-Host "########################################"
	Write-Host ""
	Write_LogEntry -Message "Anzahl Treiber mit Updates: $($driversToUpdate.Count)" -Level "INFO"

	# Output the drivers with newer versions
	#if ($driversToUpdate.Count -gt 0) {
		#$driversToUpdate | Format-Table -AutoSize
		#$driversToUpdate | Sort-Object -Property DriverName | Format-Table -Property DriverName, Version, DirectoryPath -AutoSize
	#} else {
		#Write-Host "Alle Treiber sind aktuell." -ForegroundColor "Green"
	#}

	############################################################################
	# Compare Online Drivers and Versions with local one End
	############################################################################

	############################################################################
	# Download Drivers into Temp Subfolder DriverDownloads Start
	############################################################################

	function ExtractExeFile {
		$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

		# Check if 7-Zip is installed by verifying the existence of the executable
		if (Test-Path $sevenZipPath) {
			#Write-Host "7-Zip ist installiert."

			# Extract the filename without extension to create the target folder
			$fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($downloadFileName)
			$targetFolder = Join-Path $InstallationFolder $fileNameWithoutExtension

			# Create the target folder if it doesn't exist
			if (-not (Test-Path $targetFolder)) {
				New-Item -Path $targetFolder -ItemType Directory > $null 2>&1
				#Write-Host "Ordner erstellt: $targetFolder"
			}

			# Use 7-Zip to extract the downloaded file into the target folder
			Write-Host "Die Datei wird extrahiert." -ForegroundColor "yellow"
			#& "$sevenZipPath" x "$downloadPath" -o"$targetFolder" -y
			& "$sevenZipPath" x "$downloadPath" -o"$targetFolder" -y *> $null

			#Write-Host "Die Datei wurde nach $targetFolder extrahiert." -ForegroundColor "yellow"
		} #else {
			#Write-Host "7-Zip ist nicht installiert. Bitte installieren Sie 7-Zip, um fortzufahren."
		#}
	}

	if ($driversToUpdate.Count -gt 0) {
		Write_LogEntry -Message "Beginne Herunterladen von $($driversToUpdate.Count) Treibern." -Level "INFO"
		# Create a temporary folder
		$tempFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "DriverDownloads")
		if (-not (Test-Path -Path $tempFolder)) {
			New-Item -Path $tempFolder -ItemType Directory | Out-Null
		}
		Write_LogEntry -Message "Download TempFolder: $($tempFolder)" -Level "DEBUG"

		# Initialize the download logic
		foreach ($driver in $driversToUpdate) {
			Write-Host ""
			Write-Host "$($driver.DriverName) wird heruntergeladen.." -ForegroundColor "green"
			Write-Host ""
			$downloadUrl = $driver.DownloadLink
			$downloadFileName = [System.IO.Path]::GetFileName($downloadUrl)  # Get the filename from the URL

			# Remove anything after .zip in the filename (including the query parameters)
			$downloadFileName = $downloadFileName -replace '\?.*$', ''

			$downloadPath = [System.IO.Path]::Combine($tempFolder, $downloadFileName)
			Write_LogEntry -Message "Starte Download $($driver.DriverName) von $($downloadUrl) nach $($downloadPath)" -Level "INFO"

			# Start downloading the driver
			#Write-Host "Downloading $($driver.DriverName) from $downloadUrl to $downloadPath"
			#Write-Host "Filename: $($downloadFileName)"

			Import-Module BitsTransfer -ErrorAction SilentlyContinue
			$useBitTransfer = $null -ne (Get-Module -Name BitsTransfer -ListAvailable)

			if ($useBitTransfer) {
				Start-BitsTransfer -Source $downloadUrl -Destination $downloadPath
				Write_LogEntry -Message "BitsTransfer abgeschlossen: $($downloadPath)" -Level "INFO"
			} else {
				$webClient = New-Object System.Net.WebClient
				[void](Invoke-DownloadFile -Url $downloadUrl -OutFile $downloadPath)
				$webClient.Dispose()
				Write_LogEntry -Message "WebClient-Download abgeschlossen: $($downloadPath)" -Level "INFO"
			}
				
			# Check if the file was completely downloaded
			if (Test-Path $downloadPath) {
				Write_LogEntry -Message "Download erfolgreich: $($downloadPath). Beginne Extraktion." -Level "INFO"
				ExtractExeFile
				Write_LogEntry -Message "Extraktion abgeschlossen für: $($downloadFileName)" -Level "SUCCESS"
				
				# Remove the downloaded zip
				Remove-Item -Path $downloadPath -Force
				Write_LogEntry -Message "Heruntergeladene Datei gelöscht: $($downloadPath)" -Level "DEBUG"
				
				# Extract the filename without extension to create the target folder
				$fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($downloadFileName)
				$targetFolder = Join-Path $InstallationFolder $fileNameWithoutExtension
				
				if($driver.DriverName -eq "Armoury Crate"){
					#Write-Host "$($driver.DirectoryPath) heruntergalden und entpackt." -ForegroundColor "Yellow"
					
					$installationPath = "$InstallationFolder\ArmouryCrateInstallTool"
					$targetFolderPattern = "ArmouryCrateInstaller_*"

					# Check if the subfolder matching the pattern exists
					$subfolder = Get-ChildItem -Path $installationPath -Directory -Filter $targetFolderPattern -ErrorAction SilentlyContinue

					if ($subfolder) {
						# If subfolder exists, move it one level up
						$targetFolder = [System.IO.Path]::Combine($InstallationFolder, $subfolder.Name)
						Move-Item -Path $subfolder.FullName -Destination $targetFolder

						# Remove the now empty ArmouryCrateInstallTool folder
						Remove-Item -Path $installationPath -Recurse -Force
						#Write-Host "Subfolder moved and empty folder removed."
						
						# Add ArmouryCrate Service Version number to installer exe
						$ArmouryInstaller = "$targetFolder\ArmouryCrateInstaller.exe"
						
						if (Test-Path $ArmouryInstaller) {
							$svcVer = ($latestDrivers |
								Where-Object DriverName -eq 'Armoury Crate & Aura Creator Installer' |
								Select-Object -ExpandProperty ServiceVersion
							)
							if ($svcVer) {
								# Where to store the metadata (same folder, same base name + .json)
								$metaFile = "$ArmouryInstaller.json"

								# Build a simple hashtable and dump it as JSON
								$meta = @{
									ServiceVersion = $svcVer
									Tool           = 'ArmouryCrateInstaller'
									StampedOn      = (Get-Date).ToString('u')
								}

								$meta | ConvertTo-Json | Set-Content -Path $metaFile
								Write_LogEntry -Message "ArmouryCrate Metadaten geschrieben: $($metaFile)" -Level "INFO"
							}
						}
					}
				}
					
				if (Test-Path $targetFolder) {
					#Write-Host "Neuer Treiber vorhanden: $targetFolder" -ForegroundColor "Yellow"
					#Write-Host "$($driver.DirectoryPath) wird entfernt.." -ForegroundColor "Yellow"
					Remove-Item -Path $driver.DirectoryPath -Force -Recurse #alter Treiber
					
					Write-Host ""
					Write-Host "$($driver.DriverName) wurde aktualisiert.." -ForegroundColor "green"
					Write_LogEntry -Message "$($driver.DriverName) erfolgreich aktualisiert. Zielordner: $($targetFolder)" -Level "SUCCESS"
				} else {
					Write-Host "Download ist fehlgeschlagen. $($driver.DriverName) wurde nicht aktualisiert." -ForegroundColor "red"
					Write_LogEntry -Message "Zielordner nicht gefunden nach Extraktion: $($targetFolder) für $($driver.DriverName)" -Level "ERROR"
				}
			} else {
				Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor "red"
				Write_LogEntry -Message "Download fehlgeschlagen für $($driver.DriverName): $($downloadUrl)" -Level "ERROR"
			}
		}
		Remove-Item -Path $tempFolder -Recurse -Force
		Write_LogEntry -Message "Temp Download Ordner entfernt: $($tempFolder)" -Level "DEBUG"
	}

	############################################################################
	# Download Drivers into Temp Subfolder DriverDownloads End
	############################################################################

	############################################################################
	# Check Installed Driver Version Start
	############################################################################
	Write_LogEntry -Message "Prüfe installierte Treiber via WMI/CIM." -Level "INFO"

	# Define the driver names and chipset registry paths
	$driverNames = @(
		"Realtek Bluetooth Adapter",
		"Realtek High Definition Audio",
		"Realtek Gaming 2.5GbE Family Controller",
		"Realtek 8852BE Wireless LAN WiFi 6 PCI-E NIC",
		"AMD Radeon(TM) Graphics"
	)

	# Define registry paths for the chipset driver
	$chipsetRegistryPaths = @(
		@{
			Path      = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\AMD_Chipset_IODrivers"
			ValueName = "DisplayVersion"
		},
		@{
			Path      = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\AMD_Chipset_IODrivers"
			ValueName = "DisplayVersion"
		}
	)

	# Initialize results array
	$InstalledDrivers = @()

	# Retrieve installed drivers
	foreach ($driverName in $driverNames) {
		$installedDriver = Get-CimInstance -ClassName Win32_PnPSignedDriver |
			Where-Object { $_.DeviceName -eq $driverName }

		if ($null -ne $installedDriver) {
			# Add result to the array
			$InstalledDrivers += [PSCustomObject]@{
				DriverName       = $driverName
				InstalledVersion = $installedDriver.DriverVersion
			}
			Write_LogEntry -Message "Installierter Treiber gefunden: $($driverName) Version: $($installedDriver.DriverVersion)" -Level "DEBUG"
		} else {
			# If driver is not found
			$InstalledDrivers += [PSCustomObject]@{
				DriverName       = $driverName
				InstalledVersion = "Not Installed"
			}
			Write_LogEntry -Message "Installierter Treiber NICHT gefunden: $($driverName)" -Level "DEBUG"
		}
	}

	# Function to get the chipset driver version
	function Get-ChipsetDriverVersion {
		foreach ($regEntry in $chipsetRegistryPaths) {
			try {
				# Read the registry key and return its value if it exists
				if (Test-Path -Path $regEntry.Path) {
					$value = Get-ItemProperty -Path $regEntry.Path -Name $regEntry.ValueName -ErrorAction SilentlyContinue |
						Select-Object -ExpandProperty $regEntry.ValueName -ErrorAction SilentlyContinue
					if ($value) {
						Write_LogEntry -Message "Chipset Registry Wert gefunden: $($regEntry.Path) -> $($value)" -Level "DEBUG"
						return $value
					}
				}
			} catch {
				# Log the error and continue to the next registry path
				Write-Warning "Failed to read registry: $($_.Exception.Message)"
				Write_LogEntry -Message "Fehler beim Lesen der Registry $($regEntry.Path): $($_)" -Level "ERROR"
				continue
			}
		}
		return $null
	}

	# Retrieve chipset driver version
	$chipsetInstalledVersion = Get-ChipsetDriverVersion

	# Add chipset driver to results
	if ($chipsetInstalledVersion) {
		$InstalledDrivers += [PSCustomObject]@{
			DriverName       = "AMD Chipset Driver Suite"
			InstalledVersion = $chipsetInstalledVersion
		}
		Write_LogEntry -Message "Chipset installiert: $($chipsetInstalledVersion)" -Level "DEBUG"
	} else {
		$InstalledDrivers += [PSCustomObject]@{
			DriverName       = "AMD Chipset Driver Suite"
			InstalledVersion = "Not Installed"
		}
		Write_LogEntry -Message "Chipset Driver Suite nicht installiert." -Level "DEBUG"
	}

	# Output results
	#$InstalledDrivers | Format-Table -AutoSize
	Write_LogEntry -Message "Ermittelte installierte Treiber: $($InstalledDrivers.Count)" -Level "INFO"

	############################################################################
	# Check Installed Driver Version End
	############################################################################

	############################################################################
	# Get Driver Infos from KrX Nas and compare them with installed Drivers Start
	############################################################################

	$driverVersions = GatherLocalDriverInfos #Call Function to get Local Driver Infos
	Write_LogEntry -Message "Lokale driverVersions: $($driverVersions.Count)" -Level "DEBUG"
	#$driverVersions | Sort-Object -Property DriverName | Format-Table -AutoSize

	# Define a mapping between InstalledDrivers and driverVersions names
	$driverNameMapping = @{
		"Realtek Bluetooth Adapter"                    = "Bluetooth Driver"
		"Realtek High Definition Audio"                = "Audio Driver"
		"Realtek Gaming 2.5GbE Family Controller"      = "LAN Driver"
		"Realtek 8852BE Wireless LAN WiFi 6 PCI-E NIC" = "Wi-Fi Driver"
		"AMD Radeon(TM) Graphics"                      = "Graphics Driver"
		"AMD Chipset Driver Suite"                     = "Chipset Driver"
	}

	# Initialize an array to store drivers that need an update
	$newDriversToUpdate = @()

	# Compare installed drivers with downloaded versions
	foreach ($installedDriver in $InstalledDrivers) {
		# Map the InstalledDriver name to the corresponding name in driverVersions
		$mappedDriverName = $driverNameMapping[$installedDriver.DriverName]

		if ($mappedDriverName) {
			# Find the corresponding driver in driverVersions
			$localDriver = $driverVersions | Where-Object { $_.DriverName -eq $mappedDriverName }

			if ($localDriver) {
				Write-Host "$($localDriver.DriverName) ist installiert." -foregroundcolor Green
				Write-Host "	Installierte Version:       $($installedDriver.InstalledVersion)" -ForegroundColor Cyan
				Write-Host "	Installationsdatei Version: $($localDriver.Versions)" -ForegroundColor Cyan
				Write_LogEntry -Message "Vergleiche installierten Treiber $($localDriver.DriverName): Installiert $($installedDriver.InstalledVersion); Download $($localDriver.Versions)" -Level "DEBUG"
					
				# Compare versions only if driver is actually installed
				if ($installedDriver.InstalledVersion -ne "Not Installed") {
					# Compare versions
					if ([Version]$installedDriver.InstalledVersion -lt [Version]$localDriver.Versions) {
						# Add the driver to the update list if the installed version is older
						$newDriversToUpdate += [PSCustomObject]@{
							DriverName        = $localDriver.DriverName
							InstalledVersion  = $installedDriver.InstalledVersion
							DownloadedVersion = $localDriver.Versions
							DirectoryPath     = $localDriver.DirectoryPath
						}
						Write-Host "		Veraltete $($localDriver.DriverName) ist installiert. Update wird gestartet." -foregroundcolor "magenta"
						#Write-Host "Treiber benötigt ein Update: $($localDriver.DriverName)" -ForegroundColor Yellow
						#Write-Host "Installierte Version: $($installedDriver.InstalledVersion)" -ForegroundColor Yellow
						#Write-Host "Heruntergeladene Version: $($localDriver.Versions)" -ForegroundColor Yellow
						#Write-Host "Pfad: $($localDriver.DirectoryPath)" -ForegroundColor Yellow
						#Write-Host ""
						Write_LogEntry -Message "Treiber benötigt Update: $($localDriver.DriverName). Installiert: $($installedDriver.InstalledVersion); Heruntergeladen: $($localDriver.Versions)" -Level "INFO"
					} else {
						Write-Host "		Installierte Version ist aktuell." -foregroundcolor "DarkGray"
						Write_LogEntry -Message "Installierte Version aktuell für $($localDriver.DriverName)." -Level "DEBUG"
					}
				} else {
					Write-Host "		Treiber ist nicht installiert." -foregroundcolor "Yellow"
					Write_LogEntry -Message "Treiber $($localDriver.DriverName) ist nicht installiert." -Level "INFO"
				}
				Write-Host 
			} else {
				Write_LogEntry -Message "Kein passender heruntergeladener Treiber gefunden für installiertes Gerät: $($installedDriver.DriverName)" -Level "WARNING"
				#Write-Host "Kein passender Treiber in den heruntergeladenen Versionen gefunden für: $($installedDriver.DriverName)" -ForegroundColor Red
			}
		} else {
			Write_LogEntry -Message "Keine Zuordnung gefunden für installiertes Gerät: $($installedDriver.DriverName)" -Level "DEBUG"
			#Write-Host "Keine Zuordnung gefunden für: $($installedDriver.DriverName)" -ForegroundColor Red
		}
	}

	Write_LogEntry -Message "Anzahl neueDriversToUpdate: $($newDriversToUpdate.Count)" -Level "INFO"

	# Output the drivers that need an update
	#if ($newDriversToUpdate.Count -gt 0) {
		#Write-Host "Treiber, die aktualisiert werden müssen:" -ForegroundColor Cyan
		#$newDriversToUpdate | Format-Table -AutoSize
	#} else {
		#Write-Host "Alle Treiber sind aktuell." -ForegroundColor Green
	#}

	############################################################################
	# Get Driver Infos from KrX Nas and compare them with installed Drivers End
	############################################################################

	# Serialize objects into strings for argument passing
	$DriversToUpdateArgs = $newDriversToUpdate | ForEach-Object {
		"$($_.DriverName)|$($_.InstalledVersion)|$($_.DownloadedVersion)|$($_.DirectoryPath)"
	}

	$DriversToUpdateArgsString = $DriversToUpdateArgs -join ','
	Write_LogEntry -Message "DriversToUpdateArgsString vorbereitet: $($DriversToUpdateArgsString)" -Level "DEBUG"

	#Install if needed
	if($InstallationFlag) {
		Write_LogEntry -Message "InstallationFlag gesetzt -> Starte Mainboard-Install-Script mit Flag." -Level "INFO"
		#& "C:\Windows\system32\WindowsPowerShell\v1.0\powershell" -ExecutionPolicy Bypass -NoLogo -NoProfile -File "$NetworkShareDaten\Prog\InstallationScripts\Installation\TreiberAmdPcMainboardInstall.ps1" -InstallationFlag
		# Run the script and pass the serialized data
		Start-Process -FilePath "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @(
			"-ExecutionPolicy Bypass",
			"-NoLogo",
			"-NoProfile",
			"-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\TreiberAmdPcMainboardInstall.ps1`"",
			"-DriversToUpdate `"$DriversToUpdateArgsString`""
			"-InstallationFlag"
		) -NoNewWindow -Wait
		Write_LogEntry -Message "Externes Installationsscript TreiberAmdPcMainboardInstall.ps1 mit -InstallationFlag aufgerufen." -Level "INFO"

		Start-Process -FilePath "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @(
			"-ExecutionPolicy Bypass",
			"-NoLogo",
			"-NoProfile",
			"-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\ArmouryCrateInstall.ps1`"",
			"-InstallationFlag"
		) -NoNewWindow -Wait
		Write_LogEntry -Message "Externes Installationsscript ArmouryCrateInstall.ps1 mit -InstallationFlag aufgerufen." -Level "INFO"
	}
	elseif($newDriversToUpdate.Count -gt 0)	{
		Write_LogEntry -Message "Starte Mainboard-Install-Script für $($newDriversToUpdate.Count) Treiber (kein InstallationFlag)." -Level "INFO"
		#& "$NetworkShareDaten\Prog\InstallationScripts\Installation\TreiberAmdPcMainboardInstall.ps1" -DriversToUpdate $DriversToUpdateArgsString

		# Run the script and pass the serialized data
		Start-Process -FilePath "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @(
			"-ExecutionPolicy Bypass",
			"-NoLogo",
			"-NoProfile",
			"-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\TreiberAmdPcMainboardInstall.ps1`"",
			"-DriversToUpdate `"$DriversToUpdateArgsString`""
		) -NoNewWindow -Wait
		Write_LogEntry -Message "Externes Installationsscript TreiberAmdPcMainboardInstall.ps1 ohne Flag aufgerufen." -Level "INFO"
	}
	if(!$InstallationFlag) {
		Write_LogEntry -Message "Starte ArmouryCrateInstall.ps1 (kein InstallationFlag)." -Level "INFO"
		Start-Process -FilePath "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @(
			"-ExecutionPolicy Bypass",
			"-NoLogo",
			"-NoProfile",
			"-File `"$NetworkShareDaten\Prog\InstallationScripts\Installation\ArmouryCrateInstall.ps1`""
		) -NoNewWindow -Wait
		Write_LogEntry -Message "Externes Installationsscript ArmouryCrateInstall.ps1 ohne Flag aufgerufen." -Level "INFO"
	}
	Write-Host ""	
	Write_LogEntry -Message "Treiber-Workflow abgeschlossen." -Level "SUCCESS"
} else {
	Write-Host ""
	Write-Host "		Treiber sind NICHT für dieses System geeignet." -ForegroundColor "Blue"
	Write-Host ""
	Write_LogEntry -Message "System $($PCName) ist nicht Zielsystem. Keine Aktionen ausgeführt." -Level "INFO"
}

# === Logger-Footer: automatisch eingefügt ===
Write_LogEntry -Message "Script beendet: Program=$($ProgramName), ScriptType=$($ScriptType)" -Level "INFO"
Finalize_LogSession
# === Ende Logger-Footer ===
