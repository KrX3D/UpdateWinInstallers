# ASUS Armoury Crate Update Checker and Installer
# This script checks if ASUS Armoury Crate is installed, compares with available version 
# on NAS, and installs newer version if available

param (
    [switch]$InstallationFlag = $false
)

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
} else {
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    pause
    exit
}

$ProgramName = "Armoury Crate"

# Define NAS path where Armoury Crate installers are stored
$InstallationFolder = Join-Path $NetworkShareDaten "Treiber\AMD_PC"

function Get-InstalledArmouryCrateVersion {
    # Check for Armoury Crate in Programs and Features (both 64 and 32 bit registry paths)
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($rp in $regPaths) {
        try {
            $armouryCrate = Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue |
                            Where-Object { $_.DisplayName -and ($_.DisplayName -like "*$ProgramName*") } |
                            Select-Object -First 1
            if ($armouryCrate) {
                return $armouryCrate.DisplayVersion
            }
        } catch {
            # ignore and continue with next registry path
        }
    }

    return $null
}

function Get-LatestAvailableVersion {
    # Pattern for your installer folders
    $pattern = 'ArmouryCrateInstaller_*'

    if (-not (Test-Path $InstallationFolder)) {
        Write-Warning "InstallationFolder does not exist: $InstallationFolder"
        return $null
    }

    # Find all matching subdirectories
    $folders = Get-ChildItem -Path $InstallationFolder -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like $pattern }

    if (-not $folders) {
        Write-Warning "No folders matching '$pattern' under $InstallationFolder"
        return $null
    }

    # Extract version from folder name (assumes ArmouryCrateInstaller_1.2.3.4)
    $choices = foreach ($f in $folders) {
        if ($f.Name -match 'ArmouryCrateInstaller_(\d+(?:\.\d+){1,3})') {
            [pscustomobject]@{
                Version = [version]$matches[1]
                Folder  = $f.FullName
            }
        }
    }

    if (-not $choices) {
        Write-Warning "No valid versioned folders under $InstallationFolder"
        return $null
    }

    # Pick the folder with the highest Version
    $best = $choices | Sort-Object Version -Descending | Select-Object -First 1

    # Build paths
    $exePath  = Join-Path $best.Folder 'ArmouryCrateInstaller.exe'
    $jsonPath = "$exePath.json"

    if (-not (Test-Path $exePath)) {
        Write-Warning "Expected EXE not found: $exePath"
        return $null
    }

    $serviceVersion = $null

    if (Test-Path $jsonPath) {
        try {
            $meta = Get-Content $jsonPath -Raw | ConvertFrom-Json
            if ($meta -and $meta.ServiceVersion) {
                $serviceVersion = $meta.ServiceVersion.ToString()
            }
        } catch {
            Write-Warning "Fehler beim Lesen der JSON: $jsonPath — $_"
        }
    } else {
        Write-Warning "Metadata JSON not found: $jsonPath"
    }

    [pscustomobject]@{
        Folder         = $best.Folder
        ExePath        = $exePath
        FolderVersion  = $best.Version.ToString()
        ServiceVersion = $serviceVersion
    }
}

function Install-NewerVersion {
    param (
        [string]$installerPath
    )

    try {
        Write-Host "Starting silent installation of newer $ProgramName version..." -ForegroundColor Yellow

        # Run installer with silent parameters
        $args = @("/SILENT","/NORESTART")
        $process = Start-Process -FilePath $installerPath -ArgumentList $args -Wait -PassThru -NoNewWindow

        if ($process.ExitCode -eq 0) {
            Write-Host "Installation completed successfully!" -ForegroundColor Green
        } else {
            Write-Host "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error during installation: $_" -ForegroundColor Red
    }
}

# Main script execution

# Get currently installed version (may be $null)
$currentVersion = Get-InstalledArmouryCrateVersion

# Get latest available version (object with ServiceVersion)
$latestVersionInfo = Get-LatestAvailableVersion

if ($null -ne $latestVersionInfo) {

    $latestVersionRaw = $latestVersionInfo.ServiceVersion
    if (-not $latestVersionRaw) {
        Write-Warning "Latest version metadata empty from $($latestVersionInfo.Folder)."
        Write-Host "Metadata has to look like:" -ForegroundColor Magenta
        Write-Host '
		{
		  "Tool": "ArmouryCrateInstaller",
		  "StampedOn": "2025-12-10 14:52:39Z",
		  "ServiceVersion": "6.1.18"
		}' -ForegroundColor Yellow
        Write-Host "Unable to determine latest available version. Please check your NAS connection." -ForegroundColor Red
        exit 1
    }

    # normalize comma to dot and trim
    $latestVersion = ($latestVersionRaw -replace ',', '.').Trim()

    if ($InstallationFlag) {
        Write-Host "Installing $ProgramName version $latestVersion" -ForegroundColor Cyan
        Install-NewerVersion -installerPath $latestVersionInfo.ExePath

        # remove "Armoury Crate Notice" shortcut if exists
        $shortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Armoury Crate Notice.lnk"
        if (Test-Path -Path $shortcut -PathType Leaf) {
            Write-Host "    Startmenüverknüpfung 'Armoury Crate Notice' wird entfernt." -ForegroundColor Cyan
            Remove-Item -Path $shortcut -Force
        }
    } else {
        # handle not installed case
        if ($currentVersion) {
            $curr = ($currentVersion -replace ',', '.').Trim()
            try {
                $currentVersionObj = [version]$curr
            } catch {
                # if parsing fails, treat as zero
                $currentVersionObj = [version]"0.0.0.0"
            }
            Write-Host "$ProgramName ist installiert." -ForegroundColor Green
            Write-Host "    Installierte Version:       $currentVersion" -ForegroundColor Cyan
            Write-Host "    Installationsdatei Version: $latestVersion" -ForegroundColor Cyan
        } else {
            $currentVersionObj = [version]"0.0.0.0"
            Write-Host "$ProgramName ist nicht installiert." -ForegroundColor Yellow
            Write-Host "    Gefundene Installationsdatei Version: $latestVersion" -ForegroundColor Cyan
        }

        try {
            $latestVersionObj = [version]$latestVersion
        } catch {
            $latestVersionObj = [version]"0.0.0.0"
        }

        if ($latestVersionObj -gt $currentVersionObj) {
            Write-Host "        Veraltete $ProgramName ist installiert. Update wird gestartet." -ForegroundColor Magenta
            Install-NewerVersion -installerPath $latestVersionInfo.ExePath
        } else {
            Write-Host "        Installierte Version ist aktuell." -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host "Unable to determine latest available version. Please check your NAS connection." -ForegroundColor Red
}