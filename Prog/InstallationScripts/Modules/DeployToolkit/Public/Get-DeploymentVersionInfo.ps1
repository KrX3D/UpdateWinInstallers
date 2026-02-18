function Get-LocalInstallerVersion {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PathPattern,
    [string]$FileNameRegex,
    [int]$RegexGroup = 1,
    [scriptblock]$Convert,
    [switch]$FallbackToProductVersion,
    [ValidateSet('Version','LastWriteTime')][string]$SelectionMode = 'Version'
  )

  if ($SelectionMode -eq 'LastWriteTime') {
    $file = Get-ChildItem -Path $PathPattern -File -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1

    if (-not $file) { return $null }

    $version = $null
    if ($FileNameRegex) {
      $m = [regex]::Match($file.Name, $FileNameRegex)
      if ($m.Success) {
        $raw = $m.Groups[$RegexGroup].Value
        $version = if ($Convert) { & $Convert $raw } else { ConvertTo-VersionSafe $raw }
      }
    }

    if (-not $version -and $FallbackToProductVersion) {
      try { $version = ConvertTo-VersionSafe ((Get-Item $file.FullName).VersionInfo.ProductVersion) } catch { }
    }

    [pscustomobject]@{ File = $file; Version = $version }
    return
  }

  Get-HighestVersionFile -PathPattern $PathPattern -FileNameRegex $FileNameRegex -RegexGroup $RegexGroup -Convert $Convert -FallbackToProductVersion:$FallbackToProductVersion
}

function Get-InstalledVersionInfo {
  [CmdletBinding(DefaultParameterSetName='Registry')]
  param(
    [Parameter(ParameterSetName='Registry', Mandatory)][string]$DisplayNameLike,
    [Parameter(ParameterSetName='File', Mandatory)][string]$FilePath,
    [string]$VersionProperty = 'DisplayVersion'
  )

  if ($PSCmdlet.ParameterSetName -eq 'File') {
    if (-not (Test-Path -LiteralPath $FilePath)) { return $null }

    try {
      $info = (Get-Item -LiteralPath $FilePath -ErrorAction Stop).VersionInfo
      $version = ConvertTo-VersionSafe $info.ProductVersion
      if (-not $version) {
        $version = [version]::new([int]$info.ProductMajorPart, [int]$info.ProductMinorPart, [int]$info.ProductBuildPart, [int]$info.ProductPrivatePart)
      }

      return [pscustomobject]@{
        Source       = 'File'
        FilePath     = $FilePath
        VersionRaw   = $info.ProductVersion
        Version      = $version
      }
    } catch {
      return $null
    }
  }

  $hit = Get-InstalledSoftware -DisplayNameLike $DisplayNameLike
  if (-not $hit) { return $null }

  $raw = $null
  if ($hit.PSObject.Properties[$VersionProperty]) {
    $raw = $hit.$VersionProperty
  }

  [pscustomobject]@{
    Source         = 'Registry'
    DisplayName    = $hit.DisplayName
    VersionRaw     = $raw
    Version        = ConvertTo-VersionSafe $raw
    RegistryPath   = $hit.PSPath
    RegistryRecord = $hit
  }
}

function Test-InstallerUpdateRequired {
  [CmdletBinding()]
  param(
    [version]$InstalledVersion,
    [version]$InstallerVersion
  )

  if (-not $InstalledVersion -or -not $InstallerVersion) {
    return $false
  }

  return ($InstalledVersion -lt $InstallerVersion)
}

function Get-VersionFromFileName {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Name,
    [Parameter(Mandatory)][string]$Regex,
    [int]$RegexGroup = 1,
    [scriptblock]$Convert
  )

  $m = [regex]::Match($Name, $Regex)
  if (-not $m.Success) { return $null }

  $raw = $m.Groups[$RegexGroup].Value
  if ($Convert) { return (& $Convert $raw) }
  return ConvertTo-VersionSafe $raw
}
