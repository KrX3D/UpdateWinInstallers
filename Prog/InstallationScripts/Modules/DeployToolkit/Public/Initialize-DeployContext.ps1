function Start-DeployContext {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ProgramName,
    [Parameter(Mandatory)][ValidateSet('Update','Install','Run')][string]$ScriptType,
    [Parameter(Mandatory)][string]$ScriptRoot
  )

  $loggerPath = Join-Path $ScriptRoot 'Modules\Logger\Logger.psm1'
  if (-not (Test-Path $loggerPath)) {
    throw "Logger.psm1 nicht gefunden: $loggerPath"
  }

  # IMPORTANT: make logger commands available to the calling script
  Import-Module -Name $loggerPath -Force -Global -ErrorAction Stop

  if (-not (Get-Command Write_LogEntry -ErrorAction SilentlyContinue)) {
    throw "Logger importiert, aber Write_LogEntry ist nicht verfügbar (Export/Import Problem)."
  }

  $logRoot = Join-Path $ScriptRoot 'Log'
  if (Get-Command Set_LoggerConfig -ErrorAction SilentlyContinue) {
    Set_LoggerConfig -LogRootPath $logRoot | Out-Null
  }
  if (Get-Command Initialize_LogSession -ErrorAction SilentlyContinue) {
    Initialize_LogSession -ProgramName $ProgramName -ScriptType $ScriptType | Out-Null
  }

  Write_LogEntry -Message "Logger geladen (global): $loggerPath" -Level "INFO"
}