param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Adobe Acrobat"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "AcroRdr*.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\AdobeDcInstallation.ps1"

# Helper: normalize [version] to exactly 3 components (avoids 3- vs 4-part comparison issues
#   where [version]"25.1.21223" (rev=-1) -lt [version]"25.1.21223.0" (rev=0) = TRUE)
function Get-Adobe3Part ([string]$Raw) {
    $v = Convert-AdobeToVersion -Value $Raw
    if (-not $v) { return $null }
    return [version]"$($v.Major).$($v.Minor).$($v.Build)"
}

# ── Local version ─────────────────────────────────────────────────────────────
$localFile       = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersionRaw = $null
$localVersionObj = $null

if ($localFile) {
    $rawPV           = Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    $localVersionObj = Get-Adobe3Part -Raw $rawPV
    $localVersionRaw = if ($localVersionObj) { "$($localVersionObj.Major).$('{0:D3}' -f $localVersionObj.Minor).$($localVersionObj.Build)" } else { $rawPV }
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersionRaw (→ $localVersionObj)" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version ────────────────────────────────────────────────────────────
# rdc.adobe.io requires the x-api-key header — the key "dc-get-adobereader-cdn"
# is the public key Adobe's own get.adobe.com download page sends with each request.
$onlineVersionRaw = $null
try {
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'

    $result = Invoke-RestMethod `
        -Uri 'https://rdc.adobe.io/reader/products?lang=mui&site=enterprise&os=Windows%2011&country=DE&nativeOs=Windows%2010&api_key=dc-get-adobereader-cdn' `
        -WebSession $session `
        -Headers @{
            'Accept'          = '*/*'
            'Accept-Language' = 'de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7'
            'Origin'          = 'https://get.adobe.com'
            'Referer'         = 'https://get.adobe.com/'
            'x-api-key'       = 'dc-get-adobereader-cdn'
        } `
        -ErrorAction Stop

    $onlineVersionRaw = $result.products.reader[0].version
    if ($onlineVersionRaw) {
        Write-DeployLog -Message "Online-Version gefunden: $onlineVersionRaw" -Level 'INFO'
    } else {
        Write-DeployLog -Message "Version-Feld leer in API-Antwort." -Level 'WARNING'
    }
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Online-Version: $_" -Level 'ERROR'
}
$onlineVersionObj = if ($onlineVersionRaw) { Get-Adobe3Part -Raw $onlineVersionRaw } else { $null }

Write-Host ""
Write-Host "Lokale Version: $localVersionRaw"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersionRaw" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $onlineVersionObj -and $localVersionObj) {
    if ($onlineVersionObj -gt $localVersionObj) {
        $digits   = Convert-AdobeVersionToDigits -Version $onlineVersionObj
        $fileName = "AcroRdrDCx64${digits}_de_DE.exe"
        $destPath = Join-Path $InstallationFolder $fileName
        $url      = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/$digits/$fileName"

        [void](Invoke-InstallerDownload `
            -Url                $url `
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
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFile) {
    Write-DeployLog -Message "Keine lokale Installationsdatei – Update-Vergleich übersprungen." -Level 'WARNING'
} elseif (-not $onlineVersionObj) {
    Write-DeployLog -Message "Online-Version konnte nicht ermittelt werden." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile       = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersionRaw = $null
$localVersionObj = $null
if ($localFile) {
    $rawPV           = Get-InstallerFileVersion -FilePath $localFile.FullName -Source ProductVersion
    $localVersionObj = Get-Adobe3Part -Raw $rawPV
    $localVersionRaw = if ($localVersionObj) { "$($localVersionObj.Major).$('{0:D3}' -f $localVersionObj.Minor).$($localVersionObj.Build)" } else { $rawPV }
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "Adobe Acrobat*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    $installedVerObj = Get-Adobe3Part -Raw $installedVersion

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersionRaw" -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersionRaw" -Level 'INFO'

    $Install = try { $installedVerObj -lt $localVersionObj } catch { $false }
    if ($Install) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
    } else {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
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