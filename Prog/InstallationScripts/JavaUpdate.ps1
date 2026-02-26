param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Java"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = Join-Path $config.InstallationFolder "Kicad"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "jdk*windows-x64_bin.msi"
$versionRegex    = 'jdk-([\d._]+)_windows-x64_bin\.msi'
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\JavaInstallation.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    # Filename: jdk-{major}_{fullversion}_windows-x64_bin.msi - use the full version part
    $m2 = [regex]::Match($localFile.Name, 'jdk-\d+_([\d.]+)_windows-x64_bin\.msi')
    $localVersion = if ($m2.Success) { $m2.Groups[1].Value } else {
        # Fallback: extract whatever is between jdk- and _windows
        $raw = [regex]::Match($localFile.Name, $versionRegex).Groups[1].Value
        if ($raw -like '*_*') { ($raw -split '_')[1] } else { $raw }
    }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
    Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
    return
}

# ── Online version ────────────────────────────────────────────────────────────
$onlineVersion = $null
$webContent    = $null

try {
    $webContent    = (Invoke-WebRequest -Uri 'https://www.oracle.com/java/technologies/downloads/' -UseBasicParsing -ErrorAction Stop).Content
    $latestRaw     = [regex]::Match($webContent, '<h3 id="java\d+">Java SE Development Kit ([\d.]+) downloads<\/h3>').Groups[1].Value
    if ($latestRaw) {
        $parts = $latestRaw -split '\.'
        while ($parts.Count -lt 3) { $parts += '0' }
        $onlineVersion = $parts -join '.'
    }
    Write-DeployLog -Message "Online-Version: $onlineVersion" -Level 'INFO'
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Online-Version: $_" -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($onlineVersion -and $webContent) {
    $doDownload = try { [version]$onlineVersion -gt [version]$localVersion } catch { $false }

    if ($doDownload) {
        Write-DeployLog -Message "Neuere Version verfügbar: $onlineVersion > $localVersion" -Level 'INFO'

        $linkRegex = 'href="(https:\/\/download\.oracle\.com\/java\/\d+\/latest\/jdk-([\d.]+)_windows-x64_bin\.msi)"'
        $linkMatch = [regex]::Match($webContent, $linkRegex)

        if ($linkMatch.Success) {
            $downloadLink = $linkMatch.Groups[1].Value
            $fileVersion  = $linkMatch.Groups[2].Value
            $fileName     = "jdk-${fileVersion}_${onlineVersion}_windows-x64_bin.msi"
            $destPath     = Join-Path $InstallationFolder $fileName

            [void](Invoke-InstallerDownload `
                -Url                $downloadLink `
                -OutFile            $destPath `
                -ConfirmDownload `
                -ReplaceOld `
                -RemoveFiles        @($localFile.FullName) `
                -KeepFiles          @($destPath) `
                -EmitHostStatus `
                -SuccessHostMessage "$ProgramName wurde aktualisiert.." `
                -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
                -SuccessLogMessage  "$ProgramName erfolgreich aktualisiert: $destPath" `
                -FailureLogMessage  "Download fehlgeschlagen: $destPath" `
                -Context            $ProgramName)
        } else {
            Write-DeployLog -Message "Download-Link konnte nicht aus der Seite extrahiert werden." -Level 'ERROR'
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null
if ($localFile) {
    $m2 = [regex]::Match($localFile.Name, 'jdk-\d+_([\d.]+)_windows-x64_bin\.msi')
    $localVersion = if ($m2.Success) { $m2.Groups[1].Value } else {
        $raw = [regex]::Match($localFile.Name, $versionRegex).Groups[1].Value
        if ($raw -like '*_*') { ($raw -split '_')[1] } else { $raw }
    }
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    # Normalise to 3 components for fair comparison
    $instVerNorm = (($installedVersion -split '\.')[0..2] -join '.')

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    $Install = try { [version]$instVerNorm -lt [version]$localVersion } catch { $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht in Registry gefunden." -Level 'INFO'
}

Write-Host ""

# ── Install if needed ─────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
