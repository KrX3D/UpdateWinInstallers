function Invoke-WebRequestCompat {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Uri,
    [string]$OutFile,
    [switch]$ReturnContent
  )

  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

  if ($OutFile) {
    try {
      Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop | Out-Null
      return $true
    } catch {
      try {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($Uri, $OutFile)
        $wc.Dispose()
        return $true
      } catch {
        return $false
      }
    }
  }

  try {
    $res = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop
    if ($ReturnContent) { return $res.Content }
    return $res
  } catch {
    return $null
  }
}

function Get-OnlineVersionFromContent {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Content,
    [Parameter(Mandatory)][string]$Regex,
    [int]$RegexGroup = 1,
    [switch]$SelectLast,
    [scriptblock]$Convert
  )

  $matches = [regex]::Matches($Content, $Regex)
  if (-not $matches -or $matches.Count -eq 0) { return $null }

  $m = if ($SelectLast) { $matches[$matches.Count - 1] } else { $matches[0] }
  $raw = $m.Groups[$RegexGroup].Value
  if ($Convert) { return (& $Convert $raw) }
  return ConvertTo-VersionSafe $raw
}

function Remove-FilesSafe {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PathPattern,
    [string[]]$ExcludeFullName
  )

  Get-ChildItem -Path $PathPattern -File -ErrorAction SilentlyContinue |
    Where-Object { -not $ExcludeFullName -or ($ExcludeFullName -notcontains $_.FullName) } |
    ForEach-Object {
      try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {}
    }
}

function Remove-PathSafe {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Path,
    [switch]$Recurse
  )

  if (-not (Test-Path -LiteralPath $Path)) { return $true }
  try {
    Remove-Item -LiteralPath $Path -Force -Recurse:$Recurse -ErrorAction Stop
    return $true
  } catch {
    return $false
  }
}

function Copy-FileSafe {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Destination,
    [switch]$Recurse
  )

  if (-not (Test-Path -LiteralPath $Source)) { return $false }

  $destDir = Split-Path -Path $Destination -Parent
  if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
  }

  try {
    Copy-Item -Path $Source -Destination $Destination -Force -Recurse:$Recurse -ErrorAction Stop
    return $true
  } catch {
    return $false
  }
}

function Invoke-InstallerFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$FilePath,
    [string]$Arguments,
    [switch]$Wait
  )

  if (-not (Test-Path -LiteralPath $FilePath)) { return $false }

  try {
    Start-Process -FilePath $FilePath -ArgumentList $Arguments -Wait:$Wait -ErrorAction Stop | Out-Null
    return $true
  } catch {
    return $false
  }
}
