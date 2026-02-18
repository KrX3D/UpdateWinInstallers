param(
  [switch]$InstallationFlag = $false
)

$ProgramName = "7-Zip"
$ScriptType  = "Update"

# --- Import DeployToolkit (shared helpers) ---
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

# --- Import Logger EXACTLY like original scripts (script scope) ---
$loggerPath = Join-Path $PSScriptRoot "Modules\Logger\Logger.psm1"
if (-not (Test-Path $loggerPath)) { throw "Logger.psm1 nicht gefunden: $loggerPath" }
Import-Module $loggerPath -Force -ErrorAction Stop

# Logger config + session
$logRoot = Join-Path $PSScriptRoot "Log"
Set_LoggerConfig -LogRootPath $logRoot | Out-Null
Initialize_LogSession -ProgramName $ProgramName -ScriptType $ScriptType | Out-Null

function Log {
  param([string]$Message, [string]$Level = "INFO")
  Write_LogEntry -Message $Message -Level $Level
}

# --- Import config EXACTLY like original scripts (script scope) ---
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Log "Lade Konfigurationsdatei von: $configPath" "INFO"

if (-not (Test-Path $configPath)) {
  Log "Konfigurationsdatei nicht gefunden: $configPath" "ERROR"
  Finalize_LogSession -FinalizeMessage "Abbruch: Config fehlt"
  exit 1
}

. $configPath

# Validate the SAME variables your original script relied on
if (-not $InstallationFolder) { throw "InstallationFolder ist leer/nicht gesetzt aus Config." }
if (-not $Serverip)           { throw "Serverip ist leer/nicht gesetzt aus Config." }
if (-not $PSHostPath)         { throw "PSHostPath ist leer/nicht gesetzt aus Config." }

Log "Config OK: InstallationFolder=$InstallationFolder | Serverip=$Serverip | PSHostPath=$PSHostPath" "DEBUG"

$skipDownload = $false
$downloadPageUrl = "https://www.7-zip.org/download.html"

Write_LogEntry -Message "Download-Seite URL: $downloadPageUrl" -Level "INFO"

# Find best local installer (highest version)
$localPattern = Join-Path $InstallationFolder "7z*-x64.exe"

$local = Get-HighestVersionFile `
  -PathPattern $localPattern `
  -FileNameRegex '7z(\d+)-x64\.exe' `
  -Convert { param($digits) Convert-7ZipDigitsToVersion $digits } `
  -FallbackToProductVersion

if ($local) {
  $localInstaller = $local.File.FullName
  $localVersion   = $local.Version
  Write_LogEntry -Message "Beste lokale Installationsdatei: $($local.File.Name) (Version: $localVersion)" -Level "SUCCESS"
} else {
  Write_LogEntry -Message "Keine lokale Installationsdatei gefunden - Download-Teil wird übersprungen, aber Registry-Check läuft weiter" -Level "ERROR"
  $skipDownload = $true
}

if (-not $skipDownload) {
  # Load download page
  Write_LogEntry -Message "Lade Download-Seite herunter..." -Level "INFO"
  try {
    $pageContent = (Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing -ErrorAction Stop).Content
    Write_LogEntry -Message "Download-Seite erfolgreich abgerufen (Größe: $($pageContent.Length) Zeichen)" -Level "SUCCESS"
  } catch {
    Write_LogEntry -Message "Fehler beim Abrufen der Download-Seite: $_" -Level "ERROR"
    $skipDownload = $true
  }

  if (-not $skipDownload) {
    $patternLink = '<A href="([^"]+-x64\.exe)">Download<\/A>'
    $patternVer  = 'a/7z(\d+)-x64\.exe'
    $patternBeta = '7-Zip (\d+\.\d+).+?\(beta\)'

    $betaMatch = [regex]::Match($pageContent, $patternBeta)
    $betaV = ConvertTo-VersionSafe $betaMatch.Groups[1].Value
    Write_LogEntry -Message "Beta-Version gefunden: $($betaMatch.Groups[1].Value)" -Level "DEBUG"

    $matches = [regex]::Matches($pageContent, $patternLink)

    $bestOnlineV = $null
    $bestOnlineLink = $null
    $bestOnlineDigits = $null

    foreach ($m in $matches) {
      $href = $m.Groups[1].Value
      $vm = [regex]::Match($href, $patternVer)
      if (-not $vm.Success) { continue }

      $digits = $vm.Groups[1].Value
      $v = Convert-7ZipDigitsToVersion $digits
      if (-not $v) { continue }

      if ($betaV -and ($v -eq $betaV)) { continue }

      if (-not $bestOnlineV -or $v -gt $bestOnlineV) {
        $bestOnlineV = $v
        $bestOnlineLink = $href
        $bestOnlineDigits = $digits
      }
    }

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -ForegroundColor Cyan
    Write-Host "Online Version: $bestOnlineV" -ForegroundColor Cyan
    Write-Host ""

    if ($bestOnlineV -and $bestOnlineV -gt $localVersion) {
      Write_LogEntry -Message "Neuere Version verfügbar - starte Download..." -Level "INFO"

      $downloadUrl = ($downloadPageUrl -replace "download\.html", $bestOnlineLink)
      $outFile     = Join-Path $InstallationFolder ("7z{0}-x64.exe" -f $bestOnlineDigits)

      if (Invoke-DownloadFile -Url $downloadUrl -OutFile $outFile) {
        # Remove old installers (keep new)
        Get-ChildItem -Path $localPattern -File -ErrorAction SilentlyContinue |
          Where-Object { $_.FullName -ne $outFile } |
          ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }

        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
        Write_LogEntry -Message "$ProgramName wurde erfolgreich aktualisiert auf Version $bestOnlineV" -Level "SUCCESS"
      }
    } else {
      Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
      Write_LogEntry -Message "Keine neuere Version verfügbar - $ProgramName ist aktuell" -Level "INFO"
    }
  }
}

Write-Host ""

# Registry check (installed vs local)
$installedRaw = Get-InstalledSoftwareVersion -DisplayNameLike "$ProgramName*"
$installedV   = ConvertTo-VersionSafe $installedRaw

# Re-evaluate local after potential download
$local = Get-HighestVersionFile `
  -PathPattern $localPattern `
  -FileNameRegex '7z(\d+)-x64\.exe' `
  -Convert { param($digits) Convert-7ZipDigitsToVersion $digits } `
  -FallbackToProductVersion

$localVersion = if ($local) { $local.Version } else { $null }

$Install = $false
if ($installedV) {
  Write-Host "$ProgramName ist installiert." -ForegroundColor Green
  Write-Host "    Installierte Version:       $installedV" -ForegroundColor Cyan
  Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor Cyan

  if ($localVersion -and $installedV -lt $localVersion) {
    Write-Host "        Veraltete Version erkannt. Update wird gestartet." -ForegroundColor Magenta
    $Install = $true
  } else {
    Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
  }
} else {
  Write_LogEntry -Message "$ProgramName wurde nicht in der Registrierung gefunden." -Level "INFO"
}

Write-Host ""

# Call install script if needed
$installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\7ZipInstall.ps1"

if ($InstallationFlag) {
  Write_LogEntry -Message "InstallationFlag gesetzt - starte Installation..." -Level "INFO"
  & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $installScript -InstallationFlag
}
elseif ($Install) {
  Write_LogEntry -Message "Update erforderlich - starte Installation..." -Level "INFO"
  & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $installScript
}
else {
  Write_LogEntry -Message "Keine Installation oder Update erforderlich" -Level "INFO"
}

Write-Host ""
Finalize_LogSession -FinalizeMessage "7-Zip Update-Script erfolgreich abgeschlossen"