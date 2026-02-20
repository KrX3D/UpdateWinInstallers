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
    [Parameter(Mandatory)][string[]]$Regex,
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
    [Parameter(Mandatory)][string[]]$Regex,
    [int]$RegexGroup = 1,
    [scriptblock]$Transform,
    [switch]$SelectLast,
    [System.Text.RegularExpressions.RegexOptions]$RegexOptions = ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline),
    [string]$Context = 'OnlineVersion'
  )

  Write-DeployLog -Message "[$Context] Abruf gestartet: $Url" -Level 'DEBUG'

  $content = Invoke-WebRequestCompat -Uri $Url -ReturnContent
  if (-not $content) {
    Write-DeployLog -Message "[$Context] Abruf fehlgeschlagen: $Url" -Level 'ERROR'
    return [pscustomobject]@{ Url = $Url; Content = $null; Version = $null }
  }

  $version = Get-OnlineVersionFromContent -Content $content -Regex $Regex -RegexGroup $RegexGroup -SelectLast:$SelectLast -RegexOptions $RegexOptions
  if ($Transform) {
    $version = & $Transform $version
  }

  if ($version) {
    Write-DeployLog -Message "[$Context] Abruf erfolgreich, Version: $version" -Level 'DEBUG'
  } else {
    Write-DeployLog -Message "[$Context] Keine passende Online-Version gefunden" -Level 'WARNING'
  }

  [pscustomobject]@{ Url = $Url; Content = $content; Version = $version }
}

function Invoke-InstallerDownload {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$OutFile,
    [System.Net.SecurityProtocolType]$SecurityProtocol,
    [string]$Context = 'Download'
  )

  Write-DeployLog -Message "[$Context] Download startet: $Url -> $OutFile" -Level 'DEBUG'

  $ok = $false
  if ($PSBoundParameters.ContainsKey('SecurityProtocol')) {
    $ok = Invoke-DownloadFile -Url $Url -OutFile $OutFile -SecurityProtocol $SecurityProtocol
  } else {
    $ok = Invoke-DownloadFile -Url $Url -OutFile $OutFile
  }

  if ($ok) {
    Write-DeployLog -Message "[$Context] Download erfolgreich: $OutFile" -Level 'SUCCESS'
  } else {
    Write-DeployLog -Message "[$Context] Download fehlgeschlagen: $OutFile" -Level 'ERROR'
  }

  return $ok
}

function Resolve-DownloadedInstaller {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$DownloadedFile,
    [string]$RemovePattern,
    [switch]$ReplaceOld,
    [string[]]$KeepFiles,
    [string]$Context = 'Download'
  )

  if (-not (Test-Path -LiteralPath $DownloadedFile)) {
    Write-DeployLog -Message "[$Context] Datei fehlt nach Download: $DownloadedFile" -Level 'ERROR'
    return $false
  }
  if ((Get-Item -LiteralPath $DownloadedFile).Length -le 0) {
    Write-DeployLog -Message "[$Context] Datei ist leer: $DownloadedFile" -Level 'ERROR'
    return $false
  }

  if ($ReplaceOld -and $RemovePattern) {
    $exclude = @($DownloadedFile)
    if ($KeepFiles) { $exclude += $KeepFiles }
    Remove-FilesSafe -PathPattern $RemovePattern -ExcludeFullName $exclude
    Write-DeployLog -Message "[$Context] Alte Installer bereinigt über Pattern: $RemovePattern" -Level 'INFO'
  }

  return $true
}

function Compare-VersionState {
  [CmdletBinding()]
  param(
    [version]$InstalledVersion,
    [version]$InstallerVersion,
    [string]$Context = 'VersionCompare'
  )

  Write-DeployLog -Message "[$Context] Installed=$InstalledVersion | Installer=$InstallerVersion" -Level 'INFO'

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
    [switch]$InstallRequired,
    [string]$Context = 'InstallDecision'
  )

  if ($InstallationFlag) {
    Write-DeployLog -Message "[$Context] InstallationFlag gesetzt, starte: $InstallScript" -Level 'INFO'
    $ok = Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $InstallScript -PassInstallationFlag
    if (-not $ok) {
      Write-DeployLog -Message "[$Context] Installationsskript konnte nicht gestartet werden: $InstallScript" -Level 'ERROR'
    }
    return $ok
  }

  if ($InstallRequired) {
    Write-DeployLog -Message "[$Context] Update erforderlich, starte: $InstallScript" -Level 'INFO'
    $ok = Invoke-InstallerScript -PSHostPath $PSHostPath -ScriptPath $InstallScript
    if (-not $ok) {
      Write-DeployLog -Message "[$Context] Installationsskript konnte nicht gestartet werden: $InstallScript" -Level 'ERROR'
    }
    return $ok
  }

  Write-DeployLog -Message "[$Context] Keine Installation erforderlich" -Level 'INFO'

  return $false
}
