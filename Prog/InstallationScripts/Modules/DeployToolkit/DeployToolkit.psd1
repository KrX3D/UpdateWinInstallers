@{
  RootModule        = 'DeployToolkit.psm1'
  ModuleVersion     = '1.0.0'
  GUID              = '3f7c4f35-9c8b-4f4a-88d7-2c6d8f0d1c01'
  Author            = 'KrX'
  CompanyName       = 'Personal'
  Description       = 'Shared helpers for update/install scripts (logger init, config import, version parsing, downloads, registry checks).'
  PowerShellVersion = '5.1'
  FunctionsToExport = @(
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
    'Convert-AdobeVersionToDigits'
  )
}
