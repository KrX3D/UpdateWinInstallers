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
  'ConvertTo-VersionSafe',
  'Convert-7ZipDigitsToVersion',
  'Convert-AdobeToVersion',
  'Convert-AdobeVersionToDigits'
)