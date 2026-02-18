function Get-HighestVersionFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PathPattern,
    [string]$FileNameRegex,
    [int]$RegexGroup = 1,

    # Convert extracted string (e.g. digits) => [version]
    [scriptblock]$Convert,

    # fallback to VersionInfo.ProductVersion if regex fails
    [switch]$FallbackToProductVersion
  )

  $files = Get-ChildItem -Path $PathPattern -File -ErrorAction SilentlyContinue
  if (-not $files) { return $null }

  $best = $null
  $bestV = $null

  foreach ($f in $files) {
    $v = $null

    if ($FileNameRegex) {
      $m = [regex]::Match($f.Name, $FileNameRegex)
      if ($m.Success) {
        $raw = $m.Groups[$RegexGroup].Value
        if ($Convert) {
          $v = & $Convert $raw
        } else {
          $v = ConvertTo-VersionSafe $raw
        }
      }
    }

    if (-not $v -and $FallbackToProductVersion) {
      try {
        $v = ConvertTo-VersionSafe ((Get-Item $f.FullName).VersionInfo.ProductVersion)
      } catch { }
    }

    if ($v -and (-not $bestV -or $v -gt $bestV)) {
      $bestV = $v
      $best  = $f
    }
  }

  if (-not $best) { return $null }

  [pscustomobject]@{
    File    = $best
    Version = $bestV
  }
}