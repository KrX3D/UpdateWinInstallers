function Get-InstallerFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PathPattern,
    [string]$FileNameRegex,
    [int]$RegexGroup = 1,
    [scriptblock]$Convert,
    [switch]$FallbackToProductVersion,
    [ValidateSet('Version','LastWriteTime')][string]$SelectionMode = 'Version'
  )

  return Get-LocalInstallerVersion -PathPattern $PathPattern -FileNameRegex $FileNameRegex -RegexGroup $RegexGroup -Convert $Convert -FallbackToProductVersion:$FallbackToProductVersion -SelectionMode $SelectionMode
}

function Get-OnlineVersion {
  [CmdletBinding(DefaultParameterSetName='Content')]
  param(
    [Parameter(ParameterSetName='Url', Mandatory)][string]$Url,
    [Parameter(ParameterSetName='Content', Mandatory)][string]$Content,
    [Parameter(Mandatory)][string]$Regex,
    [int]$RegexGroup = 1,
    [switch]$SelectLast,
    [scriptblock]$Convert
  )

  $sourceContent = $Content
  if ($PSCmdlet.ParameterSetName -eq 'Url') {
    $sourceContent = Invoke-WebRequestCompat -Uri $Url -ReturnContent
  }

  if (-not $sourceContent) { return $null }

  return Get-OnlineVersionFromContent -Content $sourceContent -Regex $Regex -RegexGroup $RegexGroup -SelectLast:$SelectLast -Convert $Convert
}

function Sync-InstallerFromOnline {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][version]$LocalVersion,
    [Parameter(Mandatory)][version]$OnlineVersion,
    [Parameter(Mandatory)][string]$DownloadUrl,
    [Parameter(Mandatory)][string]$TargetFile,
    [string]$CleanupPattern,
    [switch]$KeepTargetOnly
  )

  if ($OnlineVersion -le $LocalVersion) {
    return [pscustomobject]@{
      Downloaded   = $false
      Updated      = $false
      TargetFile   = $TargetFile
      OnlineNewer  = $false
    }
  }

  $downloaded = Invoke-DownloadFile -Url $DownloadUrl -OutFile $TargetFile
  if (-not $downloaded) {
    return [pscustomobject]@{
      Downloaded   = $false
      Updated      = $false
      TargetFile   = $TargetFile
      OnlineNewer  = $true
    }
  }

  if ($CleanupPattern -and $KeepTargetOnly) {
    Remove-FilesSafe -PathPattern $CleanupPattern -ExcludeFullName @($TargetFile)
  }

  [pscustomobject]@{
    Downloaded   = $true
    Updated      = $true
    TargetFile   = $TargetFile
    OnlineNewer  = $true
  }
}

function Get-RegistryVersion {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$DisplayNameLike,
    [string]$VersionProperty = 'DisplayVersion'
  )

  return Get-InstalledVersionInfo -DisplayNameLike $DisplayNameLike -VersionProperty $VersionProperty
}

function Get-InstallerExecutionPlan {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$ProgramName,
    [version]$InstalledVersion,
    [version]$InstallerVersion,
    [switch]$InstallationFlag
  )

  $shouldInstall = $false
  $reason = 'Keine Aktion erforderlich'

  if ($InstallationFlag) {
    $shouldInstall = $true
    $reason = 'InstallationFlag gesetzt'
  }
  elseif ($InstalledVersion -and $InstallerVersion -and (Test-InstallerUpdateRequired -InstalledVersion $InstalledVersion -InstallerVersion $InstallerVersion)) {
    $shouldInstall = $true
    $reason = "Update erforderlich: Installiert=$InstalledVersion, Installer=$InstallerVersion"
  }

  [pscustomobject]@{
    ProgramName        = $ProgramName
    InstalledVersion   = $InstalledVersion
    InstallerVersion   = $InstallerVersion
    InstallationFlag   = [bool]$InstallationFlag
    ShouldExecute      = $shouldInstall
    PassInstallationFlag = [bool]$InstallationFlag
    Reason             = $reason
  }
}

function Invoke-InstallerScript {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PSHostPath,
    [Parameter(Mandatory)][string]$ScriptPath,
    [switch]$PassInstallationFlag
  )

  if (-not (Test-Path -LiteralPath $ScriptPath)) { return $false }

  $args = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
  if ($PassInstallationFlag) {
    $args += '-InstallationFlag'
  }

  try {
    & $PSHostPath @args
    return $true
  }
  catch {
    return $false
  }
}

function Invoke-ProgramUninstall {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$UninstallFile,
    [string]$Arguments,
    [switch]$Wait
  )

  return Invoke-InstallerFile -FilePath $UninstallFile -Arguments $Arguments -Wait:$Wait
}

function Invoke-ProgramInstallFromPattern {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PathPattern,
    [string]$Arguments,
    [switch]$Wait,
    [switch]$SelectNewest
  )

  $files = Get-ChildItem -Path $PathPattern -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
  if (-not $files) { return @() }
  if ($SelectNewest) { $files = @($files | Select-Object -First 1) }

  $installed = New-Object System.Collections.Generic.List[string]
  foreach ($file in $files) {
    if (Invoke-InstallerFile -FilePath $file.FullName -Arguments $Arguments -Wait:$Wait) {
      $installed.Add($file.FullName)
    }
  }

  return $installed.ToArray()
}

function Remove-StartMenuEntries {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string[]]$Paths
  )

  foreach ($path in $Paths) {
    if (Test-Path -LiteralPath $path) {
      Remove-PathSafe -Path $path -Recurse | Out-Null
    }
  }
}

function Set-UserFileAssociations {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$SetUserFtaPath,
    [Parameter(Mandatory)][string]$ApplicationPath,
    [Parameter(Mandatory)][string[]]$Extensions
  )

  if (-not (Test-Path -LiteralPath $SetUserFtaPath)) { return $false }

  foreach ($extension in ($Extensions | Sort-Object -Unique)) {
    & $SetUserFtaPath --reg $ApplicationPath $extension
  }

  return $true
}

function Get-OnlineVersionInfo {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$Regex,
    [int]$RegexGroup = 1,
    [scriptblock]$Transform,
    [switch]$SelectLast
  )

  $content = Invoke-WebRequestCompat -Uri $Url -ReturnContent
  if (-not $content) {
    return [pscustomobject]@{ Url = $Url; Content = $null; Version = $null }
  }

  $version = Get-OnlineVersionFromContent -Content $content -Regex $Regex -RegexGroup $RegexGroup -SelectLast:$SelectLast
  if ($Transform) {
    $version = & $Transform $version
  }

  [pscustomobject]@{ Url = $Url; Content = $content; Version = $version }
}

function Invoke-InstallerDownload {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$OutFile,
    [System.Net.SecurityProtocolType]$SecurityProtocol
  )

  if ($PSBoundParameters.ContainsKey('SecurityProtocol')) {
    return Invoke-DownloadFile -Url $Url -OutFile $OutFile -SecurityProtocol $SecurityProtocol
  }

  return Invoke-DownloadFile -Url $Url -OutFile $OutFile
}

function Resolve-DownloadedInstaller {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$DownloadedFile,
    [string]$RemovePattern,
    [switch]$ReplaceOld,
    [string[]]$KeepFiles
  )

  if (-not (Test-Path -LiteralPath $DownloadedFile)) { return $false }
  if ((Get-Item -LiteralPath $DownloadedFile).Length -le 0) { return $false }

  if ($ReplaceOld -and $RemovePattern) {
    $exclude = @($DownloadedFile)
    if ($KeepFiles) { $exclude += $KeepFiles }
    Remove-FilesSafe -PathPattern $RemovePattern -ExcludeFullName $exclude
  }

  return $true
}

function Compare-VersionState {
  [CmdletBinding()]
  param(
    [version]$InstalledVersion,
    [version]$InstallerVersion
  )

  [pscustomobject]@{
    InstalledVersion = $InstalledVersion
    InstallerVersion = $InstallerVersion
    IsInstalled      = [bool]$InstalledVersion
    UpdateRequired   = [bool]($InstalledVersion -and $InstallerVersion -and ($InstalledVersion -lt $InstallerVersion))
  }
}

function Invoke-InstallDecision {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PSHostPath,
    [Parameter(Mandatory)][string]$InstallScript,
    [switch]$InstallationFlag,
    [switch]$InstallRequired
  )

  if ($InstallationFlag) {
    return Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $InstallScript -PassInstallationFlag
  }

  if ($InstallRequired) {
    return Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $InstallScript
  }

  return $false
}
