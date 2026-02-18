function Invoke-DownloadFile {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$OutFile,
    [int]$Retries = 3,
    [int]$RetryDelaySeconds = 2
  )

  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

  $outDir = Split-Path -Path $OutFile -Parent
  if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }

  for ($i=1; $i -le $Retries; $i++) {
    try {
      if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "Download [$i/$Retries]: $Url -> $OutFile" -Level "INFO"
      }

      # Prefer Invoke-WebRequest; fallback to WebClient if needed
      try {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
      } catch {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($Url, $OutFile)
        $wc.Dispose()
      }

      if (Test-Path $OutFile -PathType Leaf -and (Get-Item $OutFile).Length -gt 0) {
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
      if ($i -lt $Retries) { Start-Sleep -Seconds $RetryDelaySeconds }
    }
  }

  if (Get-Command Write_LogEntry -ErrorAction SilentlyContinue) {
    Write_LogEntry -Message "Download endgültig fehlgeschlagen: $Url" -Level "ERROR"
  }
  return $false
}