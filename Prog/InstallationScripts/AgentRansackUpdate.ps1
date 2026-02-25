param(
    [switch]$InstallationFlag = $false
)

$ProgramName = "Agent Ransack"
$ScriptType  = "Update"

$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (-not (Test-Path $dtPath)) { throw "DeployToolkit fehlt: $dtPath" }
Import-Module $dtPath -Force -ErrorAction Stop

Start-DeployContext -ProgramName $ProgramName -ScriptType $ScriptType -ScriptRoot $PSScriptRoot

$config            = Get-DeployConfigOrExit -ScriptRoot $PSScriptRoot -ProgramName $ProgramName -FinalizeMessage "$ProgramName - Script beendet"
$InstallationFolder = $config.InstallationFolder
$Serverip           = $config.Serverip
$PSHostPath         = $config.PSHostPath

$localFileWildcard = "agentransack*.msi"
$onlineVersionUrl  = "https://www.mythicsoft.com/agentransack/"
$installScript     = "$Serverip\Daten\Prog\InstallationScripts\Installation\AgentRansackInstall.ps1"

# ── Local version ─────────────────────────────────────────────────────────────
# Version comes from the filename (build number, e.g. "3084")
$localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard |
                 Select-Object -First 1 -ExpandProperty FullName

$localVersion = if ($localFilePath) {
    [System.IO.Path]::GetFileNameWithoutExtension($localFilePath) `
        -replace 'agentransack_', '' -replace 'x64_', ''
} else { $null }

Write-DeployLog -Message "Lokale Datei: $localFilePath | Version: $localVersion" -Level 'DEBUG'

# ── Online version ────────────────────────────────────────────────────────────
# Returns raw build-number string, e.g. "3084"
$onlineInfo   = Get-OnlineVersionInfo -Url $onlineVersionUrl -Regex @('agentransack_(\d+)') -Context $ProgramName
$onlineVersion = $onlineInfo.Version   # raw string via internal -ReturnRaw

# ── Compare and download ──────────────────────────────────────────────────────
if ($localFilePath -and $onlineVersion) {
    Write-Host ""
    Write-Host "Lokale Version: $localVersion" -ForegroundColor Cyan
    Write-Host "Online Version: $onlineVersion" -ForegroundColor Cyan
    Write-Host ""

    if ([int]$onlineVersion -gt [int]$localVersion) {
        $downloadUrl = "https://download.mythicsoft.com/flp/$onlineVersion/agentransack_x64_msi_$onlineVersion.zip"

        $ok = Invoke-ZipInstallerUpdate `
            -Url         $downloadUrl `
            -ExtractTo   $InstallationFolder `
            -RemoveFiles @($localFilePath) `
            -Context     $ProgramName

        if ($ok) {
            Write-Host "$ProgramName wurde aktualisiert.." -ForegroundColor Green
            Write-DeployLog -Message "$ProgramName aktualisiert auf Version $onlineVersion." -Level 'SUCCESS'
        } else {
            Write-Host "Download ist fehlgeschlagen. $ProgramName wurde nicht aktualisiert." -ForegroundColor Red
        }
    } else {
        Write-Host "Kein Online Update verfügbar. $ProgramName ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Kein Update erforderlich." -Level 'INFO'
    }
} elseif (-not $localFilePath) {
    Write-DeployLog -Message "Keine lokale Installationsdatei gefunden." -Level 'WARNING'
}

Write-Host ""

# ── Re-evaluate local file after potential update ─────────────────────────────
$localFilePath = Get-ChildItem -Path $InstallationFolder -Filter $localFileWildcard |
                 Select-Object -First 1 -ExpandProperty FullName
$localVersion  = if ($localFilePath) {
    [System.IO.Path]::GetFileNameWithoutExtension($localFilePath) `
        -replace 'agentransack_', '' -replace 'x64_', ''
} else { $null }

# ── Installed version (build-number taken from part [2] of DisplayVersion) ────
$installedInfo = Get-RegistryVersion -DisplayNameLike "$ProgramName*"
$Install       = $false

if ($installedInfo) {
    # DisplayVersion is like "2.1.3084"; take the 3rd segment as the build number
    $installedVersion = ($installedInfo.VersionRaw -split '\.')[2]

    Write-Host "$ProgramName ist installiert." -ForegroundColor Green
    Write-Host "    Installierte Version:       $installedVersion" -ForegroundColor Cyan
    Write-Host "    Installationsdatei Version: $localVersion" -ForegroundColor Cyan
    Write-DeployLog -Message "Installiert: $installedVersion | Lokal: $localVersion" -Level 'INFO'

    if ($installedVersion -lt $localVersion) {
        Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
        Write-DeployLog -Message "Update erforderlich." -Level 'INFO'
        $Install = $true
    } elseif ($installedVersion -eq $localVersion) {
        Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        Write-DeployLog -Message "Keine Aktion erforderlich." -Level 'INFO'
    }
} else {
    Write-DeployLog -Message "$ProgramName nicht in Registry gefunden." -Level 'INFO'
}

Write-Host ""

# ── Install if needed ─────────────────────────────────────────────────────────
if ($InstallationFlag) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript -PassInstallationFlag
} elseif ($Install) {
    Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $installScript
}

Write-Host ""
Stop-DeployContext -FinalizeMessage "$ProgramName - Script beendet"
