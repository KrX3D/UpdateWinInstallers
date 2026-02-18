Set-StrictMode -Version Latest

$Public = Join-Path $PSScriptRoot 'Public'
Get-ChildItem -Path $Public -Filter '*.ps1' -File -ErrorAction SilentlyContinue | ForEach-Object {
  . $_.FullName
}

Export-ModuleMember -Function @(
  'Initialize-DeployContext',
  'Import-SharedConfig',
  'Invoke-DownloadFile',
  'Get-InstalledSoftware',
  'Get-InstalledSoftwareVersion',
  'Get-HighestVersionFile',
  'Get-LocalInstallerVersion',
  'Get-InstalledVersionInfo',
  'Test-InstallerUpdateRequired',
  'Get-VersionFromFileName',
  'Invoke-WebRequestCompat',
  'Get-OnlineVersionFromContent',
  'Remove-FilesSafe',
  'Remove-PathSafe',
  'Copy-FileSafe',
  'Invoke-InstallerFile',
  'ConvertTo-VersionSafe',
  'Convert-7ZipDigitsToVersion',
  'Convert-AdobeToVersion',
  'Convert-AdobeVersionToDigits'
)
