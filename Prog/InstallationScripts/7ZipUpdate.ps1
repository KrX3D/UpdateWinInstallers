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
Log "Lade Konfigurationsdatei von: $configPath"

try {
  $config = Import-SharedConfig -ConfigPath $configPath
  $InstallationFolder = $config.InstallationFolder
  $Serverip = $config.Serverip
  $PSHostPath = $config.PSHostPath
} catch {
  Log "Konfigurationsdatei konnte nicht geladen werden: $_" "ERROR"
  Finalize_LogSession -FinalizeMessage "Abbruch: Config fehlt"
  exit 1
}

$localPattern = Join-Path $InstallationFolder "7z*-x64.exe"
$versionRegex = '7z(\d+)-x64\.exe'
$convert7Zip = { param($digits) Convert-7ZipDigitsToVersion $digits }

$local = Get-InstallerFile -PathPattern $localPattern -FileNameRegex $versionRegex -Convert $convert7Zip -FallbackToProductVersion
if (-not $local) {
  Log "Keine lokale Installationsdatei gefunden - Download kann nicht ausgeführt werden" "WARNING"
} else {
  Log "Beste lokale Installationsdatei: $($local.File.Name) (Version: $($local.Version))" "SUCCESS"

  $downloadPageUrl = "https://www.7-zip.org/download.html"
  $pageContent = Invoke-WebRequestCompat -Uri $downloadPageUrl -ReturnContent

  if ($pageContent) {
    $patternLink = '<A href="([^"]+-x64\.exe)">Download<\/A>'
    $patternVer  = 'a/7z(\d+)-x64\.exe'
    $patternBeta = '7-Zip (\d+\.\d+).+?\(beta\)'

    $betaVersion = Get-OnlineVersionFromContent -Content $pageContent -Regex $patternBeta
    $matches = [regex]::Matches($pageContent, $patternLink)

    $bestOnlineV = $null
    $bestOnlineLink = $null
    $bestOnlineDigits = $null

    foreach ($match in $matches) {
      $href = $match.Groups[1].Value
      $v = Get-VersionFromFileName -Name $href -Regex $patternVer -Convert $convert7Zip
      if (-not $v) { continue }
      if ($betaVersion -and ($v -eq $betaVersion)) { continue }

      if (-not $bestOnlineV -or $v -gt $bestOnlineV) {
        $bestOnlineV = $v
        $bestOnlineLink = $href
        $bestOnlineDigits = [regex]::Match($href, $patternVer).Groups[1].Value
      }
    }

    Write-Host ""
    Write-Host "Lokale Version: $($local.Version)" -ForegroundColor Cyan
    Write-Host "Online Version: $bestOnlineV" -ForegroundColor Cyan
    Write-Host ""

    if ($bestOnlineV) {
      $downloadUrl = ($downloadPageUrl -replace "download\.html", $bestOnlineLink)
      $outFile     = Join-Path $InstallationFolder ("7z{0}-x64.exe" -f $bestOnlineDigits)

      $sync = Sync-InstallerFromOnline -LocalVersion $local.Version -OnlineVersion $bestOnlineV -DownloadUrl $downloadUrl -TargetFile $outFile -CleanupPattern $localPattern -KeepTargetOnly

      if ($sync.Updated) {
        Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
        Log "$ProgramName wurde erfolgreich aktualisiert auf Version $bestOnlineV" "SUCCESS"
      } elseif (-not $sync.OnlineNewer) {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Log "Keine neuere Version verfügbar - $ProgramName ist aktuell"
      } else {
        Log "Neue Version erkannt, Download aber fehlgeschlagen" "ERROR"
      }
    }
  } else {
    Log "Download-Seite konnte nicht geladen werden" "ERROR"
  }
}

Write-Host ""

$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedV = if ($installedInfo) { $installedInfo.Version } else { $null }
$localAfter = Get-InstallerFile -PathPattern $localPattern -FileNameRegex $versionRegex -Convert $convert7Zip -FallbackToProductVersion
$localVersion = if ($localAfter) { $localAfter.Version } else { $null }

if ($installedV) {
  Write-Host "$ProgramName ist installiert." -ForegroundColor Green
  Write-Host "    Installierte Version:       $installedV" -ForegroundColor Cyan
  Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor Cyan
} else {
  Log "$ProgramName wurde nicht in der Registrierung gefunden."
}

$plan = Get-InstallerExecutionPlan -ProgramName $ProgramName -InstalledVersion $installedV -InstallerVersion $localVersion -InstallationFlag:$InstallationFlag
if ($plan.ShouldExecute) {
  if ($plan.PassInstallationFlag) {
    Write-Host "        InstallationFlag gesetzt. Installation wird gestartet." -ForegroundColor Magenta
  } else {
    Write-Host "        Veraltete Version erkannt. Update wird gestartet." -ForegroundColor Magenta
  }

  $installScript = "$Serverip\Daten\Prog\InstallationScripts\Installation\7ZipInstall.ps1"
  $started = Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag:$plan.PassInstallationFlag
  if (-not $started) {
    Log "Installationsskript konnte nicht gestartet werden: $installScript" "ERROR"
  }
} else {
  Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
  Log "Keine Installation oder Update erforderlich ($($plan.Reason))"
}

Write-Host ""
Finalize_LogSession -FinalizeMessage "7-Zip Update-Script erfolgreich abgeschlossen"
