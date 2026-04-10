function Get-GitHubLatestRelease {
  <#
  .SYNOPSIS
    Queries the GitHub API for the latest release of a repository and returns
    asset download info.
  .OUTPUTS
    PSCustomObject: Tag, Version, DownloadUrl, AssetName, AllAssets, Release
    Returns $null on failure.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Repo,
    [string]$Token,
    [string]$AssetNamePattern,
    [scriptblock]$AssetFilter,
    [string]$VersionRegex = '(\d+\.\d+\.\d+)',
    [int]$VersionGroup = 1,
    [int]$Retries = 3,
    [int]$RetryDelaySeconds = 2,
    [string]$Context = 'GitHubRelease'
  )

  $apiUrl  = "https://api.github.com/repos/$Repo/releases/latest"
  $headers = @{
    'User-Agent' = 'DeployToolkit/1.0'
    'Accept'     = 'application/vnd.github.v3+json'
  }
  if ($Token) { $headers['Authorization'] = "token $Token" }

  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

  Write-DeployLog -Message "[$Context] GitHub API: $apiUrl" -Level 'DEBUG'

  $release = $null
  for ($i = 1; $i -le $Retries; $i++) {
    try {
      $release = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
      break
    } catch {
      Write-DeployLog -Message "[$Context] GitHub API Versuch $i/$Retries fehlgeschlagen: $($_.Exception.Message)" -Level 'WARNING'
      if ($i -lt $Retries) { Start-Sleep -Seconds $RetryDelaySeconds }
    }
  }

  if (-not $release) {
    Write-DeployLog -Message "[$Context] GitHub API endgültig fehlgeschlagen nach $Retries Versuchen." -Level 'ERROR'
    return $null
  }

  $asset = if ($AssetFilter) {
    $release.assets | Where-Object { & $AssetFilter $_ } | Select-Object -First 1
  } elseif ($AssetNamePattern) {
    $release.assets | Where-Object { $_.name -match $AssetNamePattern } | Select-Object -First 1
  } else {
    $release.assets | Select-Object -First 1
  }

  $downloadUrl = if ($asset) { $asset.browser_download_url } else { $null }
  $assetName   = if ($asset) { $asset.name } else { $null }

  $version = $null
  $sources = @()
  if ($assetName) { $sources += $assetName }
  $sources += ($release.tag_name -replace '^v', '')

  foreach ($src in $sources) {
    $m = [regex]::Match($src, $VersionRegex)
    if ($m.Success) { $version = $m.Groups[$VersionGroup].Value; break }
  }

  Write-DeployLog -Message "[$Context] Asset=$assetName | Version=$version | URL=$downloadUrl" -Level 'DEBUG'

  [pscustomobject]@{
    Tag         = $release.tag_name
    Version     = $version
    DownloadUrl = $downloadUrl
    AssetName   = $assetName
    AllAssets   = $release.assets
    Release     = $release
  }
}