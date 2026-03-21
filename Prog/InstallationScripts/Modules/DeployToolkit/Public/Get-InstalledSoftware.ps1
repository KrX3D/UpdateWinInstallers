function Get-InstalledSoftware {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$DisplayNameLike)

  $paths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
  )

  Write-DeployLog -Message "Durchsuche Registry-Pfade nach installierten Programmen..." -Level 'DEBUG'

  $results = foreach ($p in $paths) {
    Write-DeployLog -Message "Prüfe Registry-Pfad: $p" -Level 'DEBUG'

    if (Test-Path $p) {
      Get-ChildItem $p -ErrorAction SilentlyContinue |
        Get-ItemProperty -ErrorAction SilentlyContinue |
        Where-Object {
          $_.PSObject.Properties['DisplayName'] -and
          $_.DisplayName -like $DisplayNameLike
        }
    }
  }

  # Sort by DisplayVersion descending so the highest installed version wins
  # when multiple registry entries match (e.g. 32-bit and 64-bit entries)
  $results |
    Where-Object { $_.PSObject.Properties['DisplayVersion'] } |
    Sort-Object { [version]($_.DisplayVersion -replace '[^\d\.]','') } -Descending -ErrorAction SilentlyContinue |
    Select-Object -First 1
}

function Get-InstalledSoftwareVersion {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$DisplayNameLike)

  $hit = Get-InstalledSoftware -DisplayNameLike $DisplayNameLike
  if ($hit -and $hit.PSObject.Properties['DisplayVersion']) { return $hit.DisplayVersion }
  return $null
}