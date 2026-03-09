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

  if ($v -match '^20(\d{2})\.(\d{3})\.(\d{5})$') {
    return [version]"$([int]$Matches[1]).$([int]$Matches[2]).$([int]$Matches[3])"
  }

  if ($v -match '^20(\d{2})(\d{3})(\d{5})$') {
    return [version]"$([int]$Matches[1]).$([int]$Matches[2]).$([int]$Matches[3])"
  }

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
  $majorValue = if ($Version.Major -ge 2000) { $Version.Major % 100 } else { $Version.Major }
  $major = "{0:D2}" -f $majorValue
  $minor = "{0:D3}" -f $Version.Minor
  $build = "{0:D5}" -f $Version.Build
  return "$major$minor$build"
}

function ConvertTo-TrimmedVersionString {
  <#
  .SYNOPSIS
    Takes the first N components of a version string and removes leading zeros
    from the last component only - matching the Advanced Port Scanner convention.
  .EXAMPLE
    ConvertTo-TrimmedVersionString "2.5.03.0" -Parts 3  => "2.5.3"
  #>
  [CmdletBinding()]
  param(
    [string]$Value,
    [int]$Parts = 3
  )

  if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

  $split = $Value.Split('.')
  $taken = for ($i = 0; $i -lt $Parts -and $i -lt $split.Count; $i++) {
    $part = $split[$i]
    if ($i -eq ($Parts - 1)) {
      # Strip leading zeros only from the last kept component
      $part -replace '^0+(\d)', '$1'
    } else {
      $part
    }
  }

  return $taken -join '.'
}
