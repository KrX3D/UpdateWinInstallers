param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "KiCad"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config             = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = Join-Path $config.InstallationFolder "Kicad"
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath
$GitHubToken        = $config.GitHubToken

$localFileFilter = "kicad*.exe"
$installScript   = "$Serverip\Daten\Prog\InstallationScripts\Installation\KicadInstallation.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
$localFile    = Get-InstallerFilePath -Directory $InstallationFolder -Filter $localFileFilter
$localVersion = $null

if ($localFile) {
    $vi           = $localFile.VersionInfo
    $raw          = if ($vi.ProductVersion) { $vi.ProductVersion } else { $vi.FileVersion }
    $localVersion = ($raw -replace '^[vV]', '') -replace '-.*$'
    Write-DeployLog -Message "Lokale Datei: $($localFile.Name) | Version: $localVersion" -Level 'DEBUG'
} else {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

# ── Online version via GitHub ─────────────────────────────────────────────────
$githubInfo = Get-GitHubLatestRelease `
    -Repo        "KiCad/kicad-source-mirror" `
    -Token       $GitHubToken `
    -AssetFilter {
        param($a)
        $a.name -match '\.exe$' -and
        $a.name -notmatch '(?i)(arm64|portable|zip|tar)' -and
        $a.name -match '(?i)(win|windows|x64|x86_64|setup|installer)'
    } `
    -Context $ProgramName

$onlineVersion = $null
$downloadUrl   = $null

if ($githubInfo) {
    $onlineVersion = ($githubInfo.Tag -replace '^[vV]', '') -replace '-.*$'
    $downloadUrl   = $githubInfo.DownloadUrl
    Write-DeployLog -Message "Online-Version: $onlineVersion | Asset: $($githubInfo.AssetName)" -Level 'INFO'
} else {
    Write-DeployLog -Message "GitHub API lieferte keine Release-Informationen." -Level 'ERROR'
}

Write-Host ""
Write-Host "Lokale Version: $localVersion"  -ForegroundColor Cyan
Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
Write-Host ""

# ── Download if newer ─────────────────────────────────────────────────────────
if ($localFile -and $onlineVersion -and $downloadUrl) {
    $doDownload = try { [version]($onlineVersion -replace '[^0-9\.]','') -gt [version]($localVersion -replace '[^0-9\.]','') } catch { $onlineVersion -ne $localVersion }

    if ($doDownload) {
        Write-DeployLog -Message "Neue Version verfügbar. Starte Download: $downloadUrl" -Level 'INFO'

        $fileName     = Split-Path $downloadUrl -Leaf
        $downloadPath = Join-Path $InstallationFolder $fileName

        [void](Invoke-InstallerDownload `
            -Url                $downloadUrl `
            -OutFile            $downloadPath `
            -ConfirmDownload `
            -ReplaceOld `
            -RemoveFiles        @($localFile.FullName) `
            -KeepFiles          @($downloadPath) `
            -EmitHostStatus `
            -SuccessHostMessage "$ProgramName wurde aktualisiert.." `
            -FailureHostMessage "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." `
            -SuccessLogMessage  "$ProgramName erfolgreich aktualisiert: $downloadPath" `
            -FailureLogMessage  "Download fehlgeschlagen: $downloadPath" `
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
    $vi           = $localFile.VersionInfo
    $raw          = if ($vi.ProductVersion) { $vi.ProductVersion } else { $vi.FileVersion }
    $localVersion = ($raw -replace '^[vV]', '') -replace '-.*$'
}

# ── Installed vs. local ───────────────────────────────────────────────────────
$installedInfo    = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$installedVersion = if ($installedInfo) { $installedInfo.VersionRaw } else { $null }
$Install          = $false

if ($installedVersion) {
    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion"     -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    $Install = try { [version]($installedVersion -replace '[^0-9\.]','') -lt [version]($localVersion -replace '[^0-9\.]','') } catch { $false }
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
