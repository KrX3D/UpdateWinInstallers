function Get-RegistryVersion {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$DisplayNameLike,
    [string]$VersionProperty = 'DisplayVersion'
  )

  return Get-InstalledVersionInfo -DisplayNameLike $DisplayNameLike -VersionProperty $VersionProperty
}

function Invoke-InstallerScript {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PSHostPath,
    [Parameter(Mandatory)][string]$ScriptPath,
    [switch]$InstallationFlag
  )

  if (-not (Test-Path -LiteralPath $PSHostPath)) {
    Write-DeployLog -Message "PSHostPath nicht gefunden: $PSHostPath" -Level 'ERROR'
    return $false
  }

  if (-not (Test-Path -LiteralPath $ScriptPath)) {
    Write-DeployLog -Message "Installationsskript nicht gefunden: $ScriptPath" -Level 'ERROR'
    return $false
  }

  $args = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
  if ($InstallationFlag) {
    $args += '-InstallationFlag'
  }

  Write-DeployLog -Message "Starte Installationsskript: $ScriptPath" -Level 'INFO'

  try {
    & $PSHostPath @args
    Write-DeployLog -Message "Installationsskript abgeschlossen: $ScriptPath" -Level 'SUCCESS'
    return $true
  }
  catch {
    Write-DeployLog -Message "Installationsskript fehlgeschlagen: $ScriptPath | Fehler: $($_.Exception.Message)" -Level 'ERROR'
    return $false
  }
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
    [Parameter(Mandatory)][string[]]$Paths,
    [string]$Context = 'Shortcuts',
    [switch]$EmitHostMessages
  )

  foreach ($path in $Paths) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }

    Write-DeployLog -Message "[$Context] Prüfe Verknüpfung: $path" -Level 'DEBUG'
    if (Test-Path -LiteralPath $path) {
      if ($EmitHostMessages) {
        Write-Host "`tVerknüpfung wird entfernt: $path" -ForegroundColor Cyan
      }

      if (Remove-PathSafe -Path $path -Recurse) {
        Write-DeployLog -Message "[$Context] Verknüpfung entfernt: $path" -Level 'SUCCESS'
      } else {
        Write-DeployLog -Message "[$Context] Konnte Verknüpfung nicht entfernen: $path" -Level 'WARNING'
      }
    } else {
      Write-DeployLog -Message "[$Context] Verknüpfung nicht gefunden: $path" -Level 'DEBUG'
    }
  }
}

function Import-RegistryFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$PSHostPath,
    [Parameter(Mandatory)][string]$RegistryImportScript,
    [Parameter(Mandatory)][string]$RegFilePath,
    [string]$Context = 'Registry'
  )

  if (-not (Test-Path -LiteralPath $PSHostPath)) {
    Write-DeployLog -Message "[$Context] PSHostPath nicht gefunden: $PSHostPath" -Level 'ERROR'
    return $false
  }
  if (-not (Test-Path -LiteralPath $RegistryImportScript)) {
    Write-DeployLog -Message "[$Context] RegistryImport Script nicht gefunden: $RegistryImportScript" -Level 'ERROR'
    return $false
  }
  if (-not (Test-Path -LiteralPath $RegFilePath)) {
    Write-DeployLog -Message "[$Context] Reg-Datei nicht gefunden: $RegFilePath" -Level 'ERROR'
    return $false
  }

  Write-DeployLog -Message "[$Context] Starte Registry-Import: $RegFilePath" -Level 'INFO'
  try {
    & $PSHostPath -NoLogo -NoProfile -ExecutionPolicy Bypass -File $RegistryImportScript -Path $RegFilePath
    Write-DeployLog -Message "[$Context] Registry-Import abgeschlossen: $RegFilePath" -Level 'SUCCESS'
    return $true
  } catch {
    Write-DeployLog -Message "[$Context] Registry-Import fehlgeschlagen: $RegFilePath. Fehler: $($_.Exception.Message)" -Level 'ERROR'
    return $false
  }
}

function Set-UserFileAssociations {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$SetUserFtaPath,
    [Parameter(Mandatory)][string]$ApplicationPath,
    [Parameter(Mandatory)][string[]]$Extensions,
    [string]$Context = 'FileAssociations'
  )

  if (-not (Test-Path -LiteralPath $SetUserFtaPath)) {
    Write-DeployLog -Message "[$Context] SFTA nicht gefunden: $SetUserFtaPath" -Level 'WARNING'
    return $false
  }

  foreach ($extension in ($Extensions | Sort-Object -Unique)) {
    Write-DeployLog -Message "[$Context] Setze Zuordnung: $extension -> $ApplicationPath" -Level 'DEBUG'
    & $SetUserFtaPath --reg $ApplicationPath $extension
  }

  Write-DeployLog -Message "[$Context] Dateizuordnungen abgeschlossen" -Level 'SUCCESS'
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

  $version = Get-OnlineVersionFromContent -Content $content -Regex $Regex -RegexGroup $RegexGroup -SelectLast:$SelectLast -ReturnRaw -RegexOptions $RegexOptions
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
    [int]$Retries = 1,
    [int]$RetryDelaySeconds = 2,
    [System.Net.SecurityProtocolType]$SecurityProtocol,
    [switch]$ConfirmDownload,
    [switch]$ReplaceOld,
    [string[]]$RemoveFiles,
    [string]$RemovePattern,
    [string[]]$KeepFiles,
    [switch]$EmitHostStatus,
    [string]$SuccessHostMessage,
    [string]$FailureHostMessage,
    [string]$SuccessLogMessage,
    [string]$FailureLogMessage,
    [string]$Context = 'Download'
  )

  Write-DeployLog -Message "[$Context] Download startet: $Url -> $OutFile" -Level 'DEBUG'

  $ok = $false
  if ($PSBoundParameters.ContainsKey('SecurityProtocol')) {
    $ok = Invoke-DownloadFile -Url $Url -OutFile $OutFile -Retries $Retries -RetryDelaySeconds $RetryDelaySeconds -SecurityProtocol $SecurityProtocol
  } else {
    $ok = Invoke-DownloadFile -Url $Url -OutFile $OutFile -Retries $Retries -RetryDelaySeconds $RetryDelaySeconds
  }

  if ($ok -and $ConfirmDownload) {
    $ok = Confirm-DownloadedInstaller -DownloadedFile $OutFile -ReplaceOld:$ReplaceOld -RemoveFiles $RemoveFiles -RemovePattern $RemovePattern -KeepFiles $KeepFiles -Context $Context
  }

  if ($ok) {
    Write-DeployLog -Message "[$Context] Download erfolgreich: $OutFile" -Level 'SUCCESS'
    if ($SuccessLogMessage) {
      Write-DeployLog -Message $SuccessLogMessage -Level 'SUCCESS'
    }
    if ($EmitHostStatus -and $SuccessHostMessage) {
      Write-Host $SuccessHostMessage -ForegroundColor Green
    }
  } else {
    Write-DeployLog -Message "[$Context] Download fehlgeschlagen: $OutFile" -Level 'ERROR'
    if ($FailureLogMessage) {
      Write-DeployLog -Message $FailureLogMessage -Level 'ERROR'
    }
    if ($EmitHostStatus -and $FailureHostMessage) {
      Write-Host $FailureHostMessage -ForegroundColor Red
    }
  }

  return $ok
}

function Confirm-DownloadedInstaller {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$DownloadedFile,
    [string]$RemovePattern,
    [switch]$ReplaceOld,
    [string[]]$RemoveFiles,
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

  if ($ReplaceOld -and $RemoveFiles) {
    foreach ($oldFile in $RemoveFiles) {
      if ([string]::IsNullOrWhiteSpace($oldFile)) { continue }
      if ($oldFile -eq $DownloadedFile) { continue }
      if ($KeepFiles -and ($KeepFiles -contains $oldFile)) { continue }
      if (-not (Test-Path -LiteralPath $oldFile)) { continue }

      try {
        Remove-Item -LiteralPath $oldFile -Force -ErrorAction Stop
        Write-DeployLog -Message "[$Context] Alter Installer gelöscht: $oldFile" -Level 'INFO'
      } catch {
        Write-DeployLog -Message "[$Context] Konnte alten Installer nicht löschen: $oldFile. Fehler: $($_.Exception.Message)" -Level 'WARNING'
      }
    }
  }

  if ($ReplaceOld -and $RemovePattern) {
    $exclude = @($DownloadedFile)
    if ($KeepFiles) { $exclude += $KeepFiles }
    Remove-FilesSafe -PathPattern $RemovePattern -ExcludeFullName $exclude
    Write-DeployLog -Message "[$Context] Alte Installer bereinigt über Pattern: $RemovePattern" -Level 'INFO'
  }

  return $true
}
