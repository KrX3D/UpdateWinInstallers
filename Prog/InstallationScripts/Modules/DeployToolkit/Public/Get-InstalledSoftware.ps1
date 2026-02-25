function Get-InstalledSoftware {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$DisplayNameLike)

  $paths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
  )

  if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Durchsuche Registry-Pfade nach installierten Programmen..." -Level "DEBUG"
  }

  $results = foreach ($p in $paths) {
    if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
      Write_LogEntry -Message "Prüfe Registry-Pfad: $p" -Level "DEBUG"
    }

    if (Test-Path $p) {
      Get-ChildItem $p -ErrorAction SilentlyContinue |
        Get-ItemProperty -ErrorAction SilentlyContinue |
        Where-Object {
          $_.PSObject.Properties['DisplayName'] -and
          $_.DisplayName -like $DisplayNameLike
        }
    }
  }

  $results | Select-Object -First 1
}

function Get-InstalledSoftwareVersion {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$DisplayNameLike)

  $hit = Get-InstalledSoftware -DisplayNameLike $DisplayNameLike
  if ($hit -and $hit.PSObject.Properties['DisplayVersion']) { return $hit.DisplayVersion }
  return $null
}
