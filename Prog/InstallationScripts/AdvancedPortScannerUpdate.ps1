param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Advanced Port Scanner"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config            = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFilePath = "$InstallationFolder\Advanced_Port_Scanner_*.exe"
$webPageUrl    = "https://www.advanced-port-scanner.com/de/"

# ── Local version ─────────────────────────────────────────────────────────────
# FileVersion from the exe, trimmed to 3 parts (leading zeros stripped from part 3)
$localFile    = Get-InstallerFilePath -PathPattern $localFilePath
$localVersion = $null
if ($localFile) {
    $rawVersion   = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    $localVersion = ConvertTo-TrimmedVersionString -Value $rawVersion -Parts 3
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'INFO'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden: $localFilePath" -Level 'WARNING'
}

# ── Online version + download link ────────────────────────────────────────────
$onlineInfo = Get-OnlineInstallerLink `
    -Url         $webPageUrl `
    -LinkRegex   '<a href="(https://download\.advanced-port-scanner\.com/download/files/Advanced_Port_Scanner_[\d\.]+\.exe)"' `
    -VersionRegex 'Advanced_Port_Scanner_(\d+\.\d+\.\d+)' `
    -VersionSource Link `
    -Context     $ProgramName

if ($onlineInfo.DownloadUrl -and $onlineInfo.Version) {
    $onlineVersion = $onlineInfo.Version
    $downloadLink  = $onlineInfo.DownloadUrl

    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -ForegroundColor Cyan
    Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
    Write-Host ""

    # String comparison preserved from original
    if ($localVersion -and ($onlineVersion -gt $localVersion)) {
        $fileName           = [System.IO.Path]::GetFileName($downloadLink)
        $downloadedFilePath = "$InstallationFolder\$fileName"

        [void](Invoke-InstallerDownload `
            -Url                 $downloadLink `
            -OutFile             $downloadedFilePath `
            -ConfirmDownload `
            -ReplaceOld `
            -RemoveFiles         @(if ($localFile) { $localFile.FullName }) `
            -KeepFiles           @($downloadedFilePath) `
            -EmitHostStatus `
            -SuccessHostMessage  "$ProgramName wurde aktualisiert.." `
            -FailureHostMessage  "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
            -SuccessLogMessage   "$ProgramName erfolgreich aktualisiert: $downloadedFilePath" `
            -FailureLogMessage   "Download fehlgeschlagen: $downloadedFilePath" `
            -Context             $ProgramName)
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "Downloadlink oder Version nicht gefunden auf: $webPageUrl" -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file after potential download ───────────────────────────
$localFile    = Get-InstallerFilePath -PathPattern $localFilePath
$localVersion = $null
if ($localFile) {
    $rawVersion   = Get-InstallerFileVersion -FilePath $localFile.FullName -Source FileVersion
    $localVersion = ConvertTo-TrimmedVersionString -Value $rawVersion -Parts 3
}

# ── Installed version vs. local installer ─────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    $Install = try { [version]$installedVersion -lt [version]$localVersion } catch { $false }
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
    Invoke-InstallerScript -PSHostPath $PSHostPath `
        -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\AdvancedPortScannerInstallation.ps1" `
        -PassInstallationFlag | Out-Null
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath `
        -ScriptPath "$Serverip\Daten\Prog\InstallationScripts\Installation\AdvancedPortScannerInstallation.ps1" | Out-Null
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
