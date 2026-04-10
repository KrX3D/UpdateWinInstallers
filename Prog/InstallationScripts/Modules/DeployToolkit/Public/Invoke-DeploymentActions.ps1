function Invoke-WebRequestCompat {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Uri,
    [string]$OutFile,
    [switch]$ReturnContent,
    [int]$Retries = 3,
    [int]$RetryDelaySeconds = 2
  )

  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

  if ($OutFile) {
    for ($i = 1; $i -le $Retries; $i++) {
      try {
        Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop | Out-Null
        return $true
      } catch {
        Write-DeployLog -Message "Invoke-WebRequestCompat OutFile Versuch $i/$Retries fehlgeschlagen: $_" -Level 'WARNING'
        if ($i -lt $Retries) { Start-Sleep -Seconds $RetryDelaySeconds }
      }
    }
    try {
      $wc = New-Object System.Net.WebClient
      $wc.DownloadFile($Uri, $OutFile)
      $wc.Dispose()
      return $true
    } catch {
      Write-DeployLog -Message "Invoke-WebRequestCompat WebClient Fallback fehlgeschlagen: $_" -Level 'ERROR'
      return $false
    }
  }

  for ($i = 1; $i -le $Retries; $i++) {
    try {
      $res = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop
      if ($ReturnContent) { return $res.Content }
      return $res
    } catch {
      Write-DeployLog -Message "Invoke-WebRequestCompat Versuch $i/$Retries fehlgeschlagen ($Uri): $_" -Level 'WARNING'
      if ($i -lt $Retries) { Start-Sleep -Seconds $RetryDelaySeconds }
    }
  }

  Write-DeployLog -Message "Invoke-WebRequestCompat endgültig fehlgeschlagen: $Uri" -Level 'ERROR'
  return $null
}

function Get-OnlineVersionFromContent {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Content,
    [Parameter(Mandatory)][string[]]$Regex,
    [int]$RegexGroup = 1,
    [switch]$SelectLast,
    [scriptblock]$Convert,
    [switch]$ReturnRaw,
    [System.Text.RegularExpressions.RegexOptions]$RegexOptions = ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
  )

  foreach ($pattern in $Regex) {
    $regMatches = [regex]::Matches($Content, $pattern, $RegexOptions)
    if (-not $regMatches -or $regMatches.Count -eq 0) { continue }

    $m = if ($SelectLast) { $regMatches[$regMatches.Count - 1] } else { $regMatches[0] }
    $raw = $m.Groups[$RegexGroup].Value
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }

    if ($Convert) { return (& $Convert $raw) }
    if ($ReturnRaw) { return $raw }
    return ConvertTo-VersionSafe $raw
  }

  return $null
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
    [object]$Arguments,
    [switch]$Wait
  )

  $resolvedFilePath = $FilePath
  if (-not (Test-Path -LiteralPath $resolvedFilePath)) {
    try {
      $cmd = Get-Command -Name $FilePath -ErrorAction Stop
      $resolvedFilePath = $cmd.Source
    } catch {
      return $false
    }
  }

  try {
    $startParams = @{ FilePath = $resolvedFilePath; Wait = [bool]$Wait; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Arguments')) {
      $startParams['ArgumentList'] = $Arguments
    }
    Start-Process @startParams | Out-Null
    return $true
  } catch {
    return $false
  }
}

function Copy-DirectoryContents {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Destination,
    [string]$Context = 'CopyConfig'
  )

  if (-not (Test-Path -LiteralPath $Source)) {
    Write-DeployLog -Message "[$Context] Quellverzeichnis nicht gefunden, übersprungen: $Source" -Level 'DEBUG'
    return $false
  }

  if (-not (Test-Path -LiteralPath $Destination)) {
    New-Item -Path $Destination -ItemType Directory -Force | Out-Null
    Write-DeployLog -Message "[$Context] Zielverzeichnis erstellt: $Destination" -Level 'DEBUG'
  }

  try {
    Get-ChildItem -LiteralPath $Source | Copy-Item -Destination $Destination -Recurse -Force -ErrorAction Stop
    Write-DeployLog -Message "[$Context] Inhalt kopiert: $Source -> $Destination" -Level 'SUCCESS'
    return $true
  } catch {
    Write-DeployLog -Message "[$Context] Fehler beim Kopieren: $Source -> $Destination | $_" -Level 'WARNING'
    return $false
  }
}
