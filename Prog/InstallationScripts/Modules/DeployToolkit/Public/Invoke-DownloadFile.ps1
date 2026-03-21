function Invoke-DownloadFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$OutFile,
    [int]$Retries = 3,
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
    } catch {
      try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
    }
  } else {
    try { [System.Net.ServicePointManager]::SecurityProtocol = $SecurityProtocol } catch {}
  }

  $outDir = Split-Path -Path $OutFile -Parent
  if (-not (Test-Path $outDir)) {
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null
  }

  for ($i = 1; $i -le $Retries; $i++) {
    try {
      Write-DeployLog -Message "Download [$i/$Retries]: $Url -> $OutFile" -Level 'INFO'

      $wc = New-Object System.Net.WebClient
      try {
        $wc.DownloadFile($Url, $OutFile)
      } finally {
        $wc.Dispose()
      }

      if ((Test-Path $OutFile -PathType Leaf) -and ((Get-Item $OutFile).Length -gt 0)) {
        Write-DeployLog -Message "Download OK ($((Get-Item $OutFile).Length) Bytes)" -Level 'SUCCESS'
        return $true
      }

      throw "Downloaded file missing or empty."
    } catch {
      Write-DeployLog -Message "Download fehlgeschlagen (Versuch $i/$Retries): $_" -Level 'WARNING'
      # Clean up partial file before next attempt
      if (Test-Path -LiteralPath $OutFile) {
        try { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue } catch {}
      }
      if ($i -lt $Retries) {
        Start-Sleep -Seconds $RetryDelaySeconds
      }
    }
  }

  Write-DeployLog -Message "Download endgültig fehlgeschlagen: $Url" -Level 'ERROR'
  return $false
}