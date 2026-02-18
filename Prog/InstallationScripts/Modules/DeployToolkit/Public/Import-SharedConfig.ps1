function Import-SharedConfig {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$ConfigPath)

  if (-not (Test-Path $ConfigPath)) {
    throw "Konfigurationsdatei nicht gefunden: $ConfigPath"
  }

  . $ConfigPath

  # return exactly the variables your original scripts use
  [pscustomobject]@{
    InstallationFolder = $InstallationFolder
    Serverip           = $Serverip
    PSHostPath         = $PSHostPath
  }
}