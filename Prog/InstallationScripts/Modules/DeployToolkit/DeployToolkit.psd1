@{
  RootModule        = 'DeployToolkit.psm1'
  ModuleVersion     = '1.0.0'
  GUID              = '3f7c4f35-9c8b-4f4a-88d7-2c6d8f0d1c01'
  Author            = 'KrX'
  CompanyName       = 'Personal'
  Description       = 'Shared helpers for update/install scripts (logger init, config import, version parsing, downloads, registry checks).'
  PowerShellVersion = '5.1'
  FunctionsToExport = @(
    'Start-DeployContext',
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
    'Import-RegistryFile',
    'Set-UserFileAssociations',
    'Get-OnlineVersionInfo',
    'Invoke-InstallerDownload',
    'Confirm-DownloadedInstaller',
    'Compare-VersionState',
    'Invoke-InstallDecision',
    'Write-DeployLog',
    'Stop-DeployContext',
    'Get-SharedConfigPath',
    'Get-DeployConfigOrExit',
    'Get-InstallerVersionForComparison',
    'Show-VersionStateSummary',
    'Import-DeployConfig'
  )
}
