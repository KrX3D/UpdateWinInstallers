function ConvertTo-VersionSafe {
  [CmdletBinding()]
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
  $clean = ($Value -replace '[^\d\.]', '')
  try { return [version]$clean } catch { return $null }
}

function Convert-7ZipDigitsToVersion {
  [CmdletBinding()]
  param([string]$Digits) # e.g. 2401 => 24.01

  if (-not $Digits -or $Digits -notmatch '^\d{4,}$') { return $null }
  $major = $Digits.Substring(0,2)
  $minor = $Digits.Substring(2,2)
  return ConvertTo-VersionSafe "$major.$minor"
}

function Convert-AdobeToVersion {
  [CmdletBinding()]
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
  $v = $Value.Trim()

  if ($v -match '^\d{8,10}$') {
    if ($v.Length -eq 10) {
      $major = [int]$v.Substring(0,2)
      $minor = [int]$v.Substring(2,3)
      $build = [int]$v.Substring(5,5)
      return [version]"$major.$minor.$build"
    }
    elseif ($v.Length -eq 9) {
      $major = [int]$v.Substring(0,2)
      $minor = [int]$v.Substring(2,2)
      $build = [int]$v.Substring(4,5)
      return [version]"$major.$minor.$build"
    }
    else {
      $major = [int]$v.Substring(0,2)
      $minor = [int]$v.Substring(2,2)
      $build = [int]$v.Substring(4)
      return [version]"$major.$minor.$build"
    }
  }

  return ConvertTo-VersionSafe $v
}

function Convert-AdobeVersionToDigits {
  [CmdletBinding()]
  param([version]$Version)

  if (-not $Version) { return $null }
  $major = "{0:D2}" -f $Version.Major
  $minor = "{0:D3}" -f $Version.Minor
  $build = "{0:D5}" -f $Version.Build
  return "$major$minor$build"
}