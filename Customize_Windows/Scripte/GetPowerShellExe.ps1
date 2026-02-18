<#
.SYNOPSIS
  Returns the full path to the "best" PowerShell host:
    • pwsh.exe (PowerShell 7/6) if installed, otherwise
    • powershell.exe (Windows PowerShell 5.1)
#>

<#
# 1) pwsh.exe on the PATH?
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwsh) {
    Write-Output $pwsh.Path
    return
}

# 2) Check the two most common install locations
foreach ($ver in 7,6) {
    $candidate = "$env:ProgramFiles\PowerShell\$ver\pwsh.exe"
    if (Test-Path $candidate) {
        Write-Output $candidate
        return
    }
}

# 3) Fallback to Windows PowerShell 5.1
$winPS = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
Write-Output $winPS
#>
<#
.SYNOPSIS
  Returns the full path to the "best" PowerShell host:
	• pwsh.exe (PowerShell 7/6) if installed (PATH, ProgramFiles, or registry), otherwise
	• powershell.exe (Windows PowerShell 5.1)
.OUTPUTS
  string - full path to the executable
#>

#Write-Host "Starting PowerShell host detection..."

# 1) pwsh.exe on the PATH?
#Write-Host "Checking pwsh.exe on PATH (Get-Command)..."
$pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwshCmd -and $pwshCmd.Path -and (Test-Path $pwshCmd.Path)) {
	#Write-Host "Found pwsh.exe on PATH: $($pwshCmd.Path)"
	Write-Output $pwshCmd.Path
	return
}

# 2) Check the two most common install locations (Program Files)
#Write-Host "Checking common ProgramFiles locations..."
foreach ($ver in 7,6) {
	$candidate = Join-Path $env:ProgramFiles "PowerShell\$ver\pwsh.exe"
	if (Test-Path $candidate) {
		#Write-Host "Found pwsh.exe at: $candidate"
		Write-Output $candidate
		return
	}
}

# 3) Check registry for PowerShell Core installed versions and pick the newest
$regBase = "HKLM:\SOFTWARE\Microsoft\PowerShellCore\InstalledVersions"
if (Test-Path $regBase) {
	#Write-Host "Inspecting registry for PowerShell Core installs at $regBase"
	$found = @()

	foreach ($key in Get-ChildItem -Path $regBase -ErrorAction SilentlyContinue) {
		try {
			$props = Get-ItemProperty -Path $key.PSPath -ErrorAction Stop
			$installLocation = $props.InstallLocation
			$semantic = $props.SemanticVersion
			if ($installLocation) {
				$exe = Join-Path $installLocation "pwsh.exe"
				if (Test-Path $exe) {
					# parse semantic version for sorting
					$verObj = $null
					try {
						if ($semantic) { $verObj = [Version]($semantic.Split("-")[0]) }
					} catch { $verObj = $null }
					$found += [PSCustomObject]@{
						Exe       = $exe
						Semantic  = $semantic
						VersionObj = $verObj
					}
					#Write-Host "Registry entry: $exe (version: $semantic)"
				}
			}
		} catch {
			#Write-Host "Failed to read registry key $($key.PSChildName): $($_.Exception.Message)"
		}
	}

	if ($found.Count -gt 0) {
		$chosen = $null
		$haveVersions = $found | Where-Object { $_.VersionObj -ne $null }
		if ($haveVersions.Count -gt 0) {
			$chosen = $haveVersions | Sort-Object -Property { $_.VersionObj } -Descending | Select-Object -First 1
		} else {
			$chosen = $found | Select-Object -First 1
		}

		if ($chosen -and (Test-Path $chosen.Exe)) {
			#Write-Host "Selected pwsh.exe from registry: $($chosen.Exe) (version: $($chosen.Semantic))"
			Write-Output $chosen.Exe
			return
		}
	}
}

# 4) Final fallback: Windows PowerShell 5.1
$winPS = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
if (Test-Path $winPS) {
	#Write-Host "Falling back to Windows PowerShell: $winPS"
	Write-Output $winPS
	return
} else {
	#Write-Host "Windows PowerShell executable not found at expected location: $winPS"
	Write-Output $null
	return
}