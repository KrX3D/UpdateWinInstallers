Set-StrictMode -Version Latest

$Public = Join-Path $PSScriptRoot 'Public'
Get-ChildItem -Path $Public -Filter '*.ps1' -File -ErrorAction SilentlyContinue | ForEach-Object {
  . $_.FullName
}

Export-ModuleMember -Function @(
  'Start-DeployContext',
  'Import-SharedConfig',
  'Invoke-DownloadFile',
  'Get-InstalledSoftware',
  'Get-InstalledVersionInfo',
  'ConvertTo-VersionSafe',
  'ConvertTo-TrimmedVersionString',
  'Convert-7ZipDigitsToVersion',
  'Convert-AdobeToVersion',
  'Convert-AdobeVersionToDigits',
  'Invoke-WebRequestCompat',
  'Get-OnlineVersionFromContent',
  'Get-OnlineInstallerLink',
  'Get-GitHubLatestRelease',
  'Remove-FilesSafe',
  'Remove-PathSafe',
  'Copy-FileSafe',
  'Copy-DirectoryContents',
  'Invoke-InstallerFile',
  'Get-InstallerFilePath',
  'Get-InstallerFileVersion',
  'Get-RegistryVersion',
  'Invoke-InstallerScript',
  'Invoke-ProgramInstallFromPattern',
  'Remove-StartMenuEntries',
  'Import-RegistryFile',
  'Set-UserFileAssociations',
  'Get-OnlineVersionInfo',
  'Invoke-InstallerDownload',
  'Confirm-DownloadedInstaller',
  'Write-DeployLog',
  'Stop-DeployContext',
  'Get-SharedConfigPath',
  'Get-DeployConfigOrExit',
  'Import-DeployConfig',
  'Write-VersionStatus',
  'Invoke-StandardInstallerDownload'
)
