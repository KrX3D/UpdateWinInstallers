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
  'ConvertTo-VersionSafe',
  'Convert-7ZipDigitsToVersion',
  'Convert-AdobeToVersion',
  'Convert-AdobeVersionToDigits',
  'Invoke-WebRequestCompat',
  'Get-OnlineVersionFromContent',
  'Remove-FilesSafe',
  'Remove-PathSafe',
  'Copy-FileSafe',
  'Invoke-InstallerFile',
  'Get-InstallerFile',
  'Get-InstallerFilePath',
  'Get-InstallerFileVersion',
  'Get-OnlineVersion',
  'Update-InstallerFromOnline',
  'Get-RegistryVersion',
  'Get-InstallerExecutionPlan',
  'Invoke-InstallerScript',
  'Invoke-ProgramUninstall',
  'Invoke-ProgramInstallFromPattern',
  'Remove-StartMenuEntries',
  'Set-UserFileAssociations',
  'Get-OnlineVersionInfo',
  'Invoke-InstallerDownload',
  'Confirm-DownloadedInstaller',
  'Compare-VersionState',
  'Invoke-InstallDecision',
  'Write-DeployLog',
  'Complete-DeployContext',
  'Get-SharedConfigPath',
  'Import-DeployConfig'
)
