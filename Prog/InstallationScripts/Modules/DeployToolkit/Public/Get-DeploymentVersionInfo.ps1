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
