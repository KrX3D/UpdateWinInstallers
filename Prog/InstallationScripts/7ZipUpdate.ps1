param(
  [switch]$InstallationFlag = $false
)

$ProgramName = "7-Zip"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip = $config.Serverip
$PSHostPath = $config.PSHostPath
Write-DeployLog -Message "Konfigurationsdatei importiert (DeployToolkit)." -Level "INFO"

$downloadPageUrl = "https://www.7-zip.org/download.html"
$localPattern = Join-Path $InstallationFolder "7z*-x64.exe"
$versionRegex = '7z(\d+)-x64\.exe'
$convert7Zip = { param($digits) Convert-7ZipDigitsToVersion $digits }

Write-DeployLog -Message "Download-Seite URL: $downloadPageUrl" -Level "INFO"
Write-DeployLog -Message "Suche nach lokaler Installationsdatei: $localPattern" -Level "INFO"

$local = Get-InstallerFile -PathPattern $localPattern -FileNameRegex $versionRegex -Convert $convert7Zip -FallbackToProductVersion
if (-not $local) {
  Write-DeployLog -Message "Keine lokale Installationsdatei gefunden - Download kann nicht ausgeführt werden" -Level "WARNING"
} else {
  Write-DeployLog -Message "Lokale Installationsdatei gefunden: $($local.File.Name)" -Level "SUCCESS"
  Write-DeployLog -Message "Lokale Version ermittelt: $($local.Version)" -Level "INFO"

  Write-DeployLog -Message "Lade Download-Seite herunter..." -Level "INFO"
  $pageContent = Invoke-WebRequestCompat -Uri $downloadPageUrl -ReturnContent

  if ($pageContent) {
    Write-DeployLog -Message "Download-Seite erfolgreich abgerufen (Größe: $($pageContent.Length) Zeichen)" -Level "SUCCESS"
    Write-DeployLog -Message "Analysiere Seiteninhalt für Download-Links..." -Level "INFO"

    $patternLink = '<A href="([^"]+-x64\.exe)">Download<\/A>'
    $patternVer  = 'a/7z(\d+)-x64\.exe'
    $patternBeta = '7-Zip (\d+\.\d+).+?\(beta\)'

    $betaVersion = Get-OnlineVersionFromContent -Content $pageContent -Regex $patternBeta
    Write-DeployLog -Message "Beta-Version gefunden: $betaVersion" -Level "DEBUG"

    $matches = [regex]::Matches($pageContent, $patternLink)
    Write-DeployLog -Message "Gefundene Download-Links: $($matches.Count)" -Level "INFO"

    $bestOnlineV = $null
    $bestOnlineLink = $null
    $bestOnlineDigits = $null

    foreach ($match in $matches) {
      $href = $match.Groups[1].Value
      Write-DeployLog -Message "Prüfe Download-Link: $href" -Level "DEBUG"

      $v = Get-VersionFromFileName -Name $href -Regex $patternVer -Convert $convert7Zip
      if (-not $v) {
        Write-DeployLog -Message "Version konnte aus Link nicht extrahiert werden" -Level "DEBUG"
        continue
      }

      Write-DeployLog -Message "Gefundene Version: $v" -Level "DEBUG"
      if ($betaVersion -and ($v -eq $betaVersion)) {
        Write-DeployLog -Message "Version $v ist Beta-Version und wird übersprungen" -Level "DEBUG"
        continue
      } else {
        Write-DeployLog -Message "Version $v ist keine Beta-Version" -Level "DEBUG"
      }

      if (-not $bestOnlineV -or $v -gt $bestOnlineV) {
        Write-DeployLog -Message "Vergleiche Versionen: Aktuell höchste ($bestOnlineV) vs. Gefundene ($v)" -Level "DEBUG"
        $bestOnlineV = $v
        $bestOnlineLink = $href
        $bestOnlineDigits = [regex]::Match($href, $patternVer).Groups[1].Value
      }
    }

    Write-Host ""
    Write-Host "Lokale Version: $($local.Version)" -ForegroundColor Cyan
    Write-Host "Online Version: $bestOnlineV" -ForegroundColor Cyan
    Write-Host ""

    Write-DeployLog -Message "Versionsvergleich - Lokal: $($local.Version), Online: $bestOnlineV" -Level "INFO"

    if ($bestOnlineV) {
      $downloadUrl = ($downloadPageUrl -replace "download\.html", $bestOnlineLink)
      $outFile = Join-Path $InstallationFolder ("7z{0}-x64.exe" -f $bestOnlineDigits)

      if ($bestOnlineV -gt $local.Version) {
        [void](Invoke-InstallerDownload -Url $downloadUrl -OutFile $outFile -ConfirmDownload -ReplaceOld -RemovePattern $localPattern -KeepFiles @($outFile) -Context $ProgramName -EmitHostStatus -SuccessHostMessage "$ProgramName wurde aktualisiert.." -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -SuccessLogMessage "$ProgramName wurde erfolgreich aktualisiert auf Version $bestOnlineV" -FailureLogMessage "Neue Version erkannt, Download aber fehlgeschlagen")
      } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine neuere Version verfügbar - $ProgramName ist aktuell" -Level "INFO"
      }
    }
  } else {
    Write-DeployLog -Message "Download-Seite konnte nicht geladen werden" -Level "ERROR"
  }
}

Write-Host ""
Write-DeployLog -Message "Prüfe installierte Version von $ProgramName..." -Level "INFO"

$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedV = if ($installedInfo) { $installedInfo.Version } else { $null }
$localAfter = Get-InstallerVersionForComparison -PathPattern $localPattern -FileNameRegex $versionRegex -Convert $convert7Zip -Context $ProgramName -Description "Aktuelle lokale Installationsdatei Version"
$localVersion = $localAfter.Version
if ($localAfter.InstallerFile) {
  Write-DeployLog -Message "Aktuelle lokale Installationsdatei: $($localAfter.InstallerFile.Name) (Version: $localVersion)" -Level "INFO"
}

if ($installedV) {
  Write-DeployLog -Message "$ProgramName ist installiert - Version: $installedV" -Level "SUCCESS"
  Write-DeployLog -Message "Vergleiche installierte Version ($installedV) mit lokaler Datei ($localVersion)" -Level "INFO"

  Write-Host "$ProgramName ist installiert." -ForegroundColor Green
  Write-Host "    Installierte Version:       $installedV" -ForegroundColor Cyan
  Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor Cyan

  if (-not $InstallationFlag) {
    $versionState = Compare-VersionState -InstalledVersion $installedV -InstallerVersion $localVersion -Context $ProgramName
    [void](Show-VersionStateSummary -State $versionState -ProgramName $ProgramName -Context $ProgramName)
  }
} else {
  Write-DeployLog -Message "$ProgramName wurde nicht in der Registrierung gefunden." -Level "INFO"
}

$plan = Get-InstallerExecutionPlan -ProgramName $ProgramName -InstalledVersion $installedV -InstallerVersion $localVersion -InstallationFlag:$InstallationFlag
if ($plan.ShouldExecute) {
  if ($plan.PassInstallationFlag) {
    Write-Host "        InstallationFlag gesetzt. Installation wird gestartet." -ForegroundColor Magenta
    Write-DeployLog -Message "InstallationFlag gesetzt - starte Installation mit -InstallationFlag" -Level "INFO"
  } elseif (-not ($installedV -and -not $InstallationFlag)) {
    Write-Host "        Veraltete Version erkannt. Update wird gestartet." -ForegroundColor Magenta
    Write-DeployLog -Message "Veraltete Version erkannt. Update wird gestartet." -Level "INFO"
  }

  $installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\7ZipInstall.ps1"
  $started = Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag:$plan.PassInstallationFlag
  if (-not $started) {
    Write-DeployLog -Message "Installationsskript konnte nicht gestartet werden: $installScript" -Level "ERROR"
  }
} else {
  if (-not $installedV) {
    Write-DeployLog -Message "Keine Installation oder Update erforderlich (Programm nicht installiert und kein InstallationFlag)." -Level "INFO"
  } else {
    Write-DeployLog -Message "Keine Installation oder Update erforderlich" -Level "INFO"
  }
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "7-Zip Update-Script erfolgreich abgeschlossen"
