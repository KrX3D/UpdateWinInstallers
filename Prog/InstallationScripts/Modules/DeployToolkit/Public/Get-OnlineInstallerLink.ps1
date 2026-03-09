function Get-OnlineInstallerLink {
  <#
  .SYNOPSIS
    Fetches a web page and extracts both a download URL and a version string via regex.
  .OUTPUTS
    PSCustomObject: Content, DownloadUrl, Version
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,

    # Regex to find the download link in the page content.
    [Parameter(Mandatory)][string]$LinkRegex,
    [int]$LinkGroup = 1,

    # Optional prefix prepended to the matched link value (e.g. for relative URLs).
    [string]$LinkPrefix,

    # Regex to extract the version. Applied to the link value or the full content.
    [string]$VersionRegex,
    [int]$VersionGroup = 1,

    # Whether to apply VersionRegex against the matched link or the full page content.
    [ValidateSet('Link', 'Content')][string]$VersionSource = 'Link',

    [string]$Context = 'OnlineInstallerLink'
  )

  Write-DeployLog -Message "[$Context] Abruf gestartet: $Url" -Level 'DEBUG'

  $content = Invoke-WebRequestCompat -Uri $Url -ReturnContent
  if (-not $content) {
    Write-DeployLog -Message "[$Context] Seite konnte nicht abgerufen werden: $Url" -Level 'ERROR'
    return [pscustomobject]@{ Content = $null; DownloadUrl = $null; Version = $null }
  }

  $linkMatch   = [regex]::Match($content, $LinkRegex)
  $downloadUrl = $null
  if ($linkMatch.Success) {
    $raw = $linkMatch.Groups[$LinkGroup].Value
    $downloadUrl = if ($LinkPrefix) { "$LinkPrefix$raw" } else { $raw }
  }

  $version = $null
  if ($VersionRegex) {
    $vSource = if ($VersionSource -eq 'Link' -and $downloadUrl) { $downloadUrl } else { $content }
    $vMatch  = [regex]::Match($vSource, $VersionRegex)
    if ($vMatch.Success) { $version = $vMatch.Groups[$VersionGroup].Value }
  }

  Write-DeployLog -Message "[$Context] DownloadUrl=$downloadUrl | Version=$version" -Level 'DEBUG'

  [pscustomobject]@{
    Content     = $content
    DownloadUrl = $downloadUrl
    Version     = $version
  }
}
