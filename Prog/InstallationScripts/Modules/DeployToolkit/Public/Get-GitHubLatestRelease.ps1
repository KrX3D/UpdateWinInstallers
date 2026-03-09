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
    # e.g. "arduino/arduino-ide"
    [Parameter(Mandatory)][string]$Repo,

    [string]$Token,

    # Simple name regex to pick an asset (e.g. '(?i)win.*x64.*\.exe$').
    [string]$AssetNamePattern,

    # Advanced scriptblock filter; receives each asset object as the first argument.
    # Takes priority over AssetNamePattern when supplied.
    [scriptblock]$AssetFilter,

    # Regex applied to the asset name (or tag_name fallback) to extract the version.
    [string]$VersionRegex = '(\d+\.\d+\.\d+)',
    [int]$VersionGroup = 1,

    [string]$Context = 'GitHubRelease'
  )

  $apiUrl  = "https://api.github.com/repos/$Repo/releases/latest"
  $headers = @{
    'User-Agent' = 'DeployToolkit/1.0'
    'Accept'     = 'application/vnd.github.v3+json'
  }
  if ($Token) { $headers['Authorization'] = "token $Token" }

  Write-DeployLog -Message "[$Context] GitHub API: $apiUrl" -Level 'DEBUG'

  try {
    $release = Invoke-RestMethod -Uri $apiUrl -Headers $headers -ErrorAction Stop
  } catch {
    Write-DeployLog -Message "[$Context] GitHub API Fehler: $($_.Exception.Message)" -Level 'ERROR'
    return $null
  }

  # Select the best asset
  $asset = if ($AssetFilter) {
    $release.assets | Where-Object { & $AssetFilter $_ } | Select-Object -First 1
  } elseif ($AssetNamePattern) {
    $release.assets | Where-Object { $_.name -match $AssetNamePattern } | Select-Object -First 1
  } else {
    $release.assets | Select-Object -First 1
  }

  $downloadUrl = if ($asset) { $asset.browser_download_url } else { $null }
  $assetName   = if ($asset) { $asset.name } else { $null }

  # Extract version from asset name, fall back to tag_name
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
