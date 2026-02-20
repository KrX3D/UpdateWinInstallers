function Get-InstallerFilePath {
  [CmdletBinding(DefaultParameterSetName='Pattern')]
  param(
    [Parameter(ParameterSetName='Pattern', Mandatory)][string]$PathPattern,
    [Parameter(ParameterSetName='Directory', Mandatory)][string]$Directory,
    [Parameter(ParameterSetName='Directory')][string]$Filter = '*',
    [Parameter(ParameterSetName='Directory')][string]$NameLike,
    [Parameter(ParameterSetName='Directory')][string]$ExcludeNameLike,
    [ValidateSet('LatestWriteTime','OldestWriteTime')][string]$Selection = 'LatestWriteTime'
  )

  $files = if ($PSCmdlet.ParameterSetName -eq 'Pattern') {
    Get-ChildItem -Path $PathPattern -File -ErrorAction SilentlyContinue
  } else {
    Get-ChildItem -Path $Directory -Filter $Filter -File -ErrorAction SilentlyContinue |
      Where-Object {
        (-not $NameLike -or $_.Name -like $NameLike) -and
        (-not $ExcludeNameLike -or $_.Name -notlike $ExcludeNameLike)
      }
  }

  if (-not $files) { return $null }

  $sorted = if ($Selection -eq 'OldestWriteTime') {
    $files | Sort-Object LastWriteTime
  } else {
    $files | Sort-Object LastWriteTime -Descending
  }

  return ($sorted | Select-Object -First 1)
}

function Get-InstallerFileVersion {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory, ValueFromPipelineByPropertyName)][Alias('FullName')][string]$FilePath,
    [string]$FileNameRegex,
    [int]$RegexGroup = 1,
    [ValidateSet('Auto','FileName','ProductVersion','FileVersion')][string]$Source = 'Auto',
    [scriptblock]$Convert
  )

  if (-not (Test-Path -LiteralPath $FilePath)) { return $null }

  $leaf = Split-Path -Path $FilePath -Leaf

  $fromName = {
    if (-not $FileNameRegex) { return $null }
    $m = [regex]::Match($leaf, $FileNameRegex)
    if (-not $m.Success) { return $null }
    $raw = $m.Groups[$RegexGroup].Value
    if ($Convert) { return (& $Convert $raw) }
    return $raw
  }

  $fromProductVersion = {
    try {
      $v = (Get-Item -LiteralPath $FilePath -ErrorAction Stop).VersionInfo.ProductVersion
      if ($Convert) { return (& $Convert $v) }
      return $v
    } catch { return $null }
  }

  $fromFileVersion = {
    try {
      $v = (Get-Item -LiteralPath $FilePath -ErrorAction Stop).VersionInfo.FileVersion
      if ($Convert) { return (& $Convert $v) }
      return $v
    } catch { return $null }
  }

  switch ($Source) {
    'FileName' { return (& $fromName) }
    'ProductVersion' { return (& $fromProductVersion) }
    'FileVersion' { return (& $fromFileVersion) }
    default {
      $candidate = & $fromName
      if ($candidate) { return $candidate }
      $candidate = & $fromProductVersion
      if ($candidate) { return $candidate }
      return (& $fromFileVersion)
    }
  }
}
