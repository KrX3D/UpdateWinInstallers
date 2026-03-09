param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "7-Zip"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileFilter = "7z*.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\7ZipInstallation.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $rawVer       = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    $localVersion = Convert-7ZipDigitsToVersion -Digits ($rawVer -replace '\.','' -replace '[^0-9]','')
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version (beta-filtered) ────────────────────────────────────────────
$pageContent   = $null
$onlineVersion = $null
$downloadUrl   = $null

try {
    $response    = Invoke-WebRequestCompat -Uri 'https://www.7-zip.org/download.html'
    $pageContent = $response.Content
} catch {
    Write-DeployLog -Message "Fehler beim Abrufen der Download-Seite: $_" -Level 'ERROR'
}

if ($pageContent) {
    $linkPattern = 'href="(/?a/7z(\d{4,})(?:-beta\d*)?-x64\.exe)"'
    $allMatches  = [regex]::Matches($pageContent, $linkPattern)

    $best = $null
    foreach ($m in $allMatches) {
        $link   = $m.Groups[1].Value
        $digits = $m.Groups[2].Value
        if ($link -match '-beta') { continue }

        $ver = Convert-7ZipDigitsToVersion -Digits $digits
        if ($ver -and (-not $best -or $ver -gt $best.Version)) {
            $best = [PSCustomObject]@{ Version = $ver; Link = $link }
        }
    }

    if ($best) {
        $onlineVersion = $best.Version
        $downloadUrl   = "https://www.7-zip.org/" + $best.Link.TrimStart('/')
        Write-DeployLog -Message "Online-Version (stabil): $onlineVersion" -Level 'INFO'
    } else {
        Write-DeployLog -Message "Keine stabile Online-Version gefunden." -Level 'WARNING'
    }
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $onlineVersion -and $downloadUrl) {
    $localVersionForCompare = if ($localVersion) { $localVersion } else { [version]'0.0' }

    if ($onlineVersion -gt $localVersionForCompare) {
        $fileName = Split-Path $downloadUrl -Leaf
        $destPath = Join-Path $InstallationFolder $fileName

        [void](Invoke-InstallerDownload `
            -Url                $downloadUrl `
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
}

Write-Host ""

# ── Re-evaluate local file ────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null
if ($localFile) {
    $rawVer       = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    $localVersion = Convert-7ZipDigitsToVersion -Digits ($rawVer -replace '\.','' -replace '[^0-9]','')
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    $installedVerObj = Convert-7ZipDigitsToVersion -Digits ($installedVersion -replace '\.','' -replace '[^0-9]','')

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVerObj" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"   -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVerObj | Lokal: $localVersion" -Level 'INFO'

    $Install = try { $installedVerObj -lt $localVersion } catch { $false }
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
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -InstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
