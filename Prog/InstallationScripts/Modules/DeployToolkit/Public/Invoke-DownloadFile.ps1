function Invoke-DownloadFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$OutFile,
    [int]$Retries = 1,
    [int]$RetryDelaySeconds = 2,
    [System.Net.SecurityProtocolType]$SecurityProtocol
  )

  if (-not $PSBoundParameters.ContainsKey('SecurityProtocol')) {
    try {
      $protocol = [System.Net.SecurityProtocolType]::Tls12
      if ([enum]::GetNames([System.Net.SecurityProtocolType]) -contains 'Tls13') {
        $protocol = $protocol -bor [System.Net.SecurityProtocolType]::Tls13
      }
      $protocol = $protocol -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
      [System.Net.ServicePointManager]::SecurityProtocol = $protocol
    }
    catch {
      try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
    }
  }
  else {
    try { [System.Net.ServicePointManager]::SecurityProtocol = $SecurityProtocol } catch {}
  }

  $outDir = Split-Path -Path $OutFile -Parent
  if (-not (Test-Path $outDir)) {
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null
  }

  for ($i = 1; $i -le $Retries; $i++) {
    try {
      if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "Download [$i/$Retries]: $Url -> $OutFile (WebClient)" -Level "INFO"
      }

      $wc = New-Object System.Net.WebClient
      try {
        $wc.DownloadFile($Url, $OutFile)
      }
      finally {
        $wc.Dispose()
      }

      if ((Test-Path $OutFile -PathType Leaf) -and ((Get-Item $OutFile).Length -gt 0)) {
        if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
          Write_LogEntry -Message "Download OK ($((Get-Item $OutFile).Length) Bytes)" -Level "SUCCESS"
        }
        return $true
      }

      throw "Downloaded file missing/empty."
    }
    catch {
      if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "Download fehlgeschlagen: $_" -Level "WARNING"
      }
      if ($i -lt $Retries) {
        Start-Sleep -Seconds $RetryDelaySeconds
      }
    }
  }

  if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Download endgültig fehlgeschlagen: $Url" -Level "ERROR"
  }
  return $false
}
