param(
  [switch]$InstallationFlag = $false
)

$ProgramName = "7-Zip"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Initialize-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

function Log {
  param([string]$Message, [string]$Level = "INFO")
  Write_LogEntry -Message $Message -Level $Level
}

$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
Log "Lade Konfigurationsdatei von: $configPath" "INFO"

try {
  $config = Import-SharedConfig -ConfigPath $configPath
  $InstallationFolder = $config.InstallationFolder
  $Serverip = $config.Serverip
  $PSHostPath = $config.PSHostPath
  Log "Konfigurationsdatei $configPath gefunden und importiert." "INFO"
} catch {
  Log "Konfigurationsdatei konnte nicht geladen werden: $_" "ERROR"
  Finalize_LogSession -FinalizeMessage "Abbruch: Config fehlt"
  exit 1
}

$downloadPageUrl = "https://www.7-zip.org/download.html"
$localPattern = Join-Path $InstallationFolder "7z*-x64.exe"
$versionRegex = '7z(\d+)-x64\.exe'
$convert7Zip = { param($digits) Convert-7ZipDigitsToVersion $digits }

Log "Download-Seite URL: $downloadPageUrl" "INFO"
Log "Suche nach lokaler Installationsdatei: $localPattern" "INFO"

$local = Get-InstallerFile -PathPattern $localPattern -FileNameRegex $versionRegex -Convert $convert7Zip -FallbackToProductVersion
if (-not $local) {
  Log "Keine lokale Installationsdatei gefunden - Download kann nicht ausgeführt werden" "WARNING"
} else {
  Log "Lokale Installationsdatei gefunden: $($local.File.Name)" "SUCCESS"
  Log "Lokale Version ermittelt: $($local.Version)" "INFO"

  Log "Lade Download-Seite herunter..." "INFO"
  $pageContent = Invoke-WebRequestCompat -Uri $downloadPageUrl -ReturnContent

  if ($pageContent) {
    Log "Download-Seite erfolgreich abgerufen (Größe: $($pageContent.Length) Zeichen)" "SUCCESS"
    Log "Analysiere Seiteninhalt für Download-Links..." "INFO"

    $patternLink = '<A href="([^"]+-x64\.exe)">Download<\/A>'
    $patternVer  = 'a/7z(\d+)-x64\.exe'
    $patternBeta = '7-Zip (\d+\.\d+).+?\(beta\)'

    $betaVersion = Get-OnlineVersionFromContent -Content $pageContent -Regex $patternBeta
    Log "Beta-Version gefunden: $betaVersion" "DEBUG"

    $matches = [regex]::Matches($pageContent, $patternLink)
    Log "Gefundene Download-Links: $($matches.Count)" "INFO"

    $bestOnlineV = $null
    $bestOnlineLink = $null
    $bestOnlineDigits = $null

    foreach ($match in $matches) {
      $href = $match.Groups[1].Value
      Log "Prüfe Download-Link: $href" "DEBUG"

      $v = Get-VersionFromFileName -Name $href -Regex $patternVer -Convert $convert7Zip
      if (-not $v) {
        Log "Version konnte aus Link nicht extrahiert werden" "DEBUG"
        continue
      }

      Log "Gefundene Version: $v" "DEBUG"
      if ($betaVersion -and ($v -eq $betaVersion)) {
        Log "Version $v ist Beta-Version und wird übersprungen" "DEBUG"
        continue
      } else {
        Log "Version $v ist keine Beta-Version" "DEBUG"
      }

      if (-not $bestOnlineV -or $v -gt $bestOnlineV) {
        Log "Vergleiche Versionen: Aktuell höchste ($bestOnlineV) vs. Gefundene ($v)" "DEBUG"
        $bestOnlineV = $v
        $bestOnlineLink = $href
        $bestOnlineDigits = [regex]::Match($href, $patternVer).Groups[1].Value
      }
    }

    Write-Host ""
    Write-Host "Lokale Version: $($local.Version)" -ForegroundColor Cyan
    Write-Host "Online Version: $bestOnlineV" -ForegroundColor Cyan
    Write-Host ""

    Log "Versionsvergleich - Lokal: $($local.Version), Online: $bestOnlineV" "INFO"

    if ($bestOnlineV) {
      $downloadUrl = ($downloadPageUrl -replace "download\.html", $bestOnlineLink)
      $outFile = Join-Path $InstallationFolder ("7z{0}-x64.exe" -f $bestOnlineDigits)

      $sync = Sync-InstallerFromOnline -LocalVersion $local.Version -OnlineVersion $bestOnlineV -DownloadUrl $downloadUrl -TargetFile $outFile -CleanupPattern $localPattern -KeepTargetOnly

      if ($sync.Updated) {
        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
        Log "$ProgramName wurde erfolgreich aktualisiert auf Version $bestOnlineV" "SUCCESS"
      } elseif (-not $sync.OnlineNewer) {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Log "Keine neuere Version verfügbar - $ProgramName ist aktuell" "INFO"
      } else {
        Log "Neue Version erkannt, Download aber fehlgeschlagen" "ERROR"
      }
    }
  } else {
    Log "Download-Seite konnte nicht geladen werden" "ERROR"
  }
}

Write-Host ""
Log "Prüfe installierte Version von $ProgramName..." "INFO"

$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedV = if ($installedInfo) { $installedInfo.Version } else { $null }
$localAfter = Get-InstallerFile -PathPattern $localPattern -FileNameRegex $versionRegex -Convert $convert7Zip -FallbackToProductVersion
$localVersion = if ($localAfter) { $localAfter.Version } else { $null }
if ($localAfter) {
  Log "Aktuelle lokale Installationsdatei: $($localAfter.File.Name) (Version: $localVersion)" "INFO"
}

if ($installedV) {
  Log "$ProgramName ist installiert - Version: $installedV" "SUCCESS"
  Log "Vergleiche installierte Version ($installedV) mit lokaler Datei ($localVersion)" "INFO"

  Write-Host "$ProgramName ist installiert." -ForegroundColor Green
  Write-Host "    Installierte Version:       $installedV" -ForegroundColor Cyan
  Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor Cyan
} else {
  Log "$ProgramName wurde nicht in der Registrierung gefunden." "INFO"
}

$plan = Get-InstallerExecutionPlan -ProgramName $ProgramName -InstalledVersion $installedV -InstallerVersion $localVersion -InstallationFlag:$InstallationFlag
if ($plan.ShouldExecute) {
  if ($plan.PassInstallationFlag) {
    Write-Host "        InstallationFlag gesetzt. Installation wird gestartet." -ForegroundColor Magenta
    Log "InstallationFlag gesetzt - starte Installation mit -InstallationFlag" "INFO"
  } else {
    Write-Host "        Veraltete Version erkannt. Update wird gestartet." -ForegroundColor Magenta
    Log "Veraltete Version erkannt. Update wird gestartet." "INFO"
  }

  $installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\7ZipInstall.ps1"
  $started = Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag:$plan.PassInstallationFlag
  if (-not $started) {
    Log "Installationsskript konnte nicht gestartet werden: $installScript" "ERROR"
  }
} else {
  Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
  Log "Installierte Version ist bereits aktuell" "INFO"
  Log "Keine Installation oder Update erforderlich" "INFO"
}

Write-Host ""
Finalize_LogSession -FinalizeMessage "7-Zip Update-Script erfolgreich abgeschlossen"
