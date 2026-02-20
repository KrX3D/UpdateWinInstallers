function Write-DeployLog {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('DEBUG','INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
  )

  if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message $Message -Level $Level
  }
}

function Complete-DeployContext {
  [CmdletBinding()]
  param(
    [string]$FinalizeMessage = 'Script beendet'
  )

  if (Get-Command -Name Finalize_LogSession -ErrorAction SilentlyContinue) {
    Finalize_LogSession -FinalizeMessage $FinalizeMessage
  } else {
    Write-DeployLog -Message $FinalizeMessage -Level 'INFO'
  }
}

function Get-SharedConfigPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ScriptRoot
  )

  return Join-Path -Path (Split-Path (Split-Path $ScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"
}

function Import-DeployConfig {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ScriptRoot
  )

  $configPath = Get-SharedConfigPath -ScriptRoot $ScriptRoot
  Write-DeployLog -Message "Lade Konfigurationsdatei: $configPath" -Level 'DEBUG'
  $config = Import-SharedConfig -ConfigPath $configPath

  Write-DeployLog -Message "Konfigurationsdatei geladen: $configPath" -Level 'INFO'
  return $config
}
