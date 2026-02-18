<#
    Enhanced Logger module
    Exports functions:
      - Initialize_LogSession
      - Write_LogEntry
      - Write_DetailedSystemInfo
      - Finalize_LogSession
      - Get_LastLogFile
      - Set_LoggerConfig
#>

# Module-scoped defaults (adjustable via Set_LoggerConfig)
$script:LogRetentionDays = 14
$script:DateTimeFormat    = "yyyy-MM-dd_HH-mm-ss"
$script:LogEntryFormat    = "yyyy-MM-dd HH:mm:ss"

# Default root logs path (fallback; kann per Set_LoggerConfig überschrieben werden)
$script:LogRootPath       = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\ScriptLogs"

# Session state
$script:CurrentLogFile    = $null
$script:SystemInfoFile    = $null
$script:CurrentProgram    = $null
$script:CurrentScriptType = $null
$script:GlobalErrorLogFile = $null

function Set_LoggerConfig {
    [CmdletBinding()]
    param(
        [int]$LogRetentionDays,
        [string]$LogRootPath
    )
    if ($PSBoundParameters.ContainsKey('LogRetentionDays')) { $script:LogRetentionDays = $LogRetentionDays }
    if ($PSBoundParameters.ContainsKey('LogRootPath')) {
        $script:LogRootPath = $LogRootPath
        # Set global error log path when LogRootPath is updated
        $script:GlobalErrorLogFile = Join-Path -Path $script:LogRootPath -ChildPath "ERROR_$(Get-Date -Format 'yyyy-MM').log"
    }
    return @{
        LogRetentionDays = $script:LogRetentionDays
        LogRootPath      = $script:LogRootPath
        GlobalErrorLogFile = $script:GlobalErrorLogFile
    }
}

function Get-PreferredIPv4 {
    [CmdletBinding()]
    param()

    try {
        $defaultRoute = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue |
            Sort-Object -Property RouteMetric, InterfaceMetric |
            Select-Object -First 1

        if ($defaultRoute) {
            $ip = Get-NetIPAddress -InterfaceIndex $defaultRoute.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.IPAddress -notmatch '^(169\.254|127\.)' } |
                Sort-Object -Property PrefixLength -Descending |
                Select-Object -ExpandProperty IPAddress -First 1
            if ($ip) { return $ip }
        }
    } catch { }

    try {
        $ipFallback = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -notmatch '^(169\.254|127\.)' } |
            Sort-Object -Property InterfaceMetric, PrefixLength -Descending |
            Select-Object -ExpandProperty IPAddress -First 1
        if ($ipFallback) { return $ipFallback }
    } catch { }

    return "N/A"
}

function Get-QuickSystemInfo {
    [CmdletBinding()]
    param()

    $info = @{}

    try {
        # Basic Info
        $info.ComputerName = $env:COMPUTERNAME
        $info.UserName = $env:USERNAME
        $info.Domain = $env:USERDOMAIN

        # OS Info
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $info.OSName = $os.Caption
            $info.OSVersion = $os.Version
            $info.OSBuild = $os.BuildNumber
            $info.OSArchitecture = $os.OSArchitecture
            $info.LastBootTime = $os.LastBootUpTime
        }

        # Hardware Info
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            $info.Manufacturer = $cs.Manufacturer
            $info.Model = $cs.Model
            $info.TotalRAM = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        }

        # CPU Info
        $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cpu) {
            $info.CPUName = $cpu.Name
            $info.CPUCores = $cpu.NumberOfCores
            $info.CPULogical = $cpu.NumberOfLogicalProcessors
        }

        # Network Info
        $info.IPAddress = Get-PreferredIPv4

        # PowerShell Info
        $info.PSVersion = $PSVersionTable.PSVersion.ToString()
        $info.PSEdition = $PSVersionTable.PSEdition

    } catch {
        Write-Warning "Error gathering system info: $_"
    }

    return $info
}

function Initialize_LogSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)][string]$ProgramName = "Unknown",
        [Parameter(Mandatory=$false)][string]$ScriptType  = "Run"  # z.B. "Update" oder "Install"
    )

    try {
        $timestamp = Get-Date -Format $script:DateTimeFormat
        $programFolder = Join-Path -Path $script:LogRootPath -ChildPath $ProgramName
        if (-not (Test-Path -Path $programFolder)) { New-Item -Path $programFolder -ItemType Directory -Force | Out-Null }

        # Ensure main log directory exists for global error log
        if (-not (Test-Path -Path $script:LogRootPath)) { New-Item -Path $script:LogRootPath -ItemType Directory -Force | Out-Null }

        $logFilename = "{0}_{1}.log" -f $ScriptType, $timestamp

        $script:CurrentLogFile    = Join-Path -Path $programFolder -ChildPath $logFilename
        $script:CurrentProgram    = $ProgramName
        $script:CurrentScriptType = $ScriptType
        $script:SystemInfoFile    = $null

        # Set global error log file (monthly rotation)
        $script:GlobalErrorLogFile = Join-Path -Path $script:LogRootPath -ChildPath "ERROR_$(Get-Date -Format 'yyyy-MM').log"

        # Gather system info
        $sysInfo = Get-QuickSystemInfo

        # Enhanced header with system information
        $separator = "=" * 80
        $miniSep = "-" * 80

        $header = @(
            $separator,
            "LOG SESSION STARTED: $(Get-Date -Format $script:LogEntryFormat)",
            $separator,
            "",
            "SCRIPT INFORMATION:",
            "  Program: $ProgramName",
            "  Script Type: $ScriptType",
            "  Log File: $script:CurrentLogFile",
            "",
            $miniSep,
            "SYSTEM OVERVIEW:",
            $miniSep,
            "  Computer: $($sysInfo.ComputerName)",
            "  User: $($sysInfo.Domain)\$($sysInfo.UserName)",
            "  IP Address: $($sysInfo.IPAddress)",
            "",
            "  OS: $($sysInfo.OSName)",
            "  Version: $($sysInfo.OSVersion) (Build $($sysInfo.OSBuild))",
            "  Architecture: $($sysInfo.OSArchitecture)",
            "  Last Boot: $($sysInfo.LastBootTime)",
            "",
            "  Hardware: $($sysInfo.Manufacturer) $($sysInfo.Model)",
            #"  Hardware: $($sysInfo.Model)",
            "  CPU: $($sysInfo.CPUName)",
            "  CPU Cores: $($sysInfo.CPUCores) physical / $($sysInfo.CPULogical) logical",
            "  RAM: $($sysInfo.TotalRAM) GB",
            "",
            "  PowerShell: $($sysInfo.PSVersion) ($($sysInfo.PSEdition))",
            $separator
        ) -join "`r`n"

        $header | Out-File -FilePath $script:CurrentLogFile -Encoding UTF8

        # Background-Job für Rotation (löscht alte Logs älter als $LogRetentionDays)
        Start-Job -ScriptBlock {
            param($root, $days)
            try {
                Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue |
                  Where-Object { ($_.LastWriteTime -lt (Get-Date).AddDays(-$days)) } |
                  ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
            } catch { }
        } -ArgumentList $script:LogRootPath, $script:LogRetentionDays | Out-Null

        Write_LogEntry -Message "Log session successfully initialized" -Level "SUCCESS"
        return $script:CurrentLogFile
    } catch {
        Write-Error "Initialize_LogSession failed: $_"
    }
}

function Write_LogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR","DEBUG")][string]$Level = "INFO",
        [string]$ProgramName,
        [string]$ScriptType
    )

    if (-not $script:CurrentLogFile) {
        Write-Warning "Log-Sitzung nicht initialisiert. Bitte zuerst Initialize_LogSession aufrufen."
        return
    }

    # override per-entry wenn gewünscht
    $p = if ($ProgramName) { $ProgramName } else { $script:CurrentProgram }
    $s = if ($ScriptType)  { $ScriptType  } else { $script:CurrentScriptType }

    $timestamp = Get-Date -Format $script:LogEntryFormat

    # Enhanced formatting with better alignment
    $levelPadded = $Level.PadRight(7)
    $programPadded = $p
    $entry = "{0} [{1}] [{2} / {3}] {4}" -f $timestamp, $levelPadded, $programPadded, $s, $Message

    # in Log schreiben
    $entry | Out-File -FilePath $script:CurrentLogFile -Encoding UTF8 -Append

    # Handle ERROR level - write to global error log
    if ($Level -eq "ERROR") {
        Write_GlobalErrorLog -Message $Message -ProgramName $p -ScriptType $s
    }
}

function Write_GlobalErrorLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$true)][string]$ProgramName,
        [Parameter(Mandatory=$true)][string]$ScriptType
    )

    try {
        if (-not $script:GlobalErrorLogFile) {
            $script:GlobalErrorLogFile = Join-Path -Path $script:LogRootPath -ChildPath "ERROR_$(Get-Date -Format 'yyyy-MM').log"
        }

        # Ensure the directory exists
        $errorLogDir = Split-Path -Path $script:GlobalErrorLogFile -Parent
        if (-not (Test-Path -Path $errorLogDir)) {
            New-Item -Path $errorLogDir -ItemType Directory -Force | Out-Null
        }

        $timestamp = Get-Date -Format $script:LogEntryFormat
        $separator = "-" * 120

        # Create structured error entry
        $errorEntry = @(
            "",
            $separator,
            "ERROR OCCURRED: $timestamp",
            "Program: $ProgramName",
            "Script Type: $ScriptType",
            "Source Log: $script:CurrentLogFile",
            "Computer: $env:COMPUTERNAME",
            "User: $env:USERNAME",
            "Message: $Message",
            $separator,
            ""
        ) -join "`r`n"

        $errorEntry | Out-File -FilePath $script:GlobalErrorLogFile -Encoding UTF8 -Append
    } catch {
        Write-Warning "Failed to write to global error log: $_"
    }
}

function Write_DetailedSystemInfo {
    [CmdletBinding()]
    param(
        [switch]$ForceWrite
    )

    if (-not $script:SystemInfoFile) {
        Write-Warning "SystemInfoFile nicht gesetzt. Bitte zuerst Initialize_LogSession aufrufen."
        return
    }

    try {
        $separator = "=" * 60
        $lines = @()
        $lines += $separator
        $lines += "SYSTEM INFORMATION"
        $lines += $separator
        $lines += "Generated: $(Get-Date -Format $script:LogEntryFormat)"
        $lines += ""

        # Basic System Info
        $lines += "BASIC INFORMATION:"
        $lines += "  Computer Name: $env:COMPUTERNAME"
        $lines += "  User: $env:USERNAME"
        $lines += "  Domain: $env:USERDOMAIN"
        $lines += "  Architecture: $env:PROCESSOR_ARCHITECTURE"
        $lines += ""

        # Operating System
        $lines += "OPERATING SYSTEM:"
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $lines += "  OS: $($os.Caption)"
            $lines += "  Version: $($os.Version)"
            $lines += "  Build: $($os.BuildNumber)"
            $lines += "  Install Date: $($os.InstallDate)"
            $lines += "  Last Boot: $($os.LastBootUpTime)"
        } catch {
            $lines += "  OS: <Query failed: $_>"
        }
        $lines += ""

        # PowerShell Info
        $lines += "POWERSHELL INFORMATION:"
        $lines += "  Version: $($PSVersionTable.PSVersion.ToString())"
        $lines += "  Edition: $($PSVersionTable.PSEdition)"
        $lines += "  Execution Policy: $(Get-ExecutionPolicy)"
        $lines += ""

        # Network Information
        $lines += "NETWORK INFORMATION:"
        try {
            $preferredIp = Get-PreferredIPv4
            if ($preferredIp -and $preferredIp -ne "N/A") {
                $lines += "  Preferred IPv4: $preferredIp"
            }
            $ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                   Where-Object { $_.IPAddress -notmatch '^(169\.254|127\.)' } |
                   Select-Object -ExpandProperty IPAddress -Unique
            if ($ips) {
                $lines += "  IP Addresses:"
                foreach ($ip in $ips) { $lines += "    - $ip" }
            } else {
                $lines += "  IP Addresses: None found"
            }
        } catch {
            $lines += "  IP Addresses: <Query failed: $_>"
        }
        $lines += ""

        # Hardware Information
        $lines += "HARDWARE INFORMATION:"
        try {
            $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($cpu) {
                $lines += "  CPU: $($cpu.Name)"
                $lines += "  Cores: $($cpu.NumberOfCores)"
                $lines += "  Logical Processors: $($cpu.NumberOfLogicalProcessors)"
            }

            $memory = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            if ($memory) {
                $totalRAM = [math]::Round($memory.TotalPhysicalMemory / 1GB, 2)
                $lines += "  Total RAM: $totalRAM GB"
            }
        } catch {
            $lines += "  Hardware Info: <Query failed: $_>"
        }
        $lines += ""

        # Sample of installed programs
        $lines += "INSTALLED PROGRAMS (Sample):"
        try {
            $installed = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
                          Where-Object { $_.DisplayName } |
                          Sort-Object DisplayName |
                          Select-Object -First 15 -Property DisplayName, DisplayVersion
            if ($installed) {
                foreach ($i in $installed) {
                    $version = if ($i.DisplayVersion) { " (v$($i.DisplayVersion))" } else { "" }
                    $lines += "  - $($i.DisplayName)$version"
                }
            } else {
                $lines += "  No programs found in registry"
            }
        } catch {
            $lines += "  <Query failed: $_>"
        }

        $lines += ""
        $lines += $separator
        $lines += ""

        $lines -join "`r`n" | Out-File -FilePath $script:SystemInfoFile -Encoding UTF8 -Append
        Write_LogEntry -Message "System information written to: $script:SystemInfoFile" -Level "INFO"
        return $script:SystemInfoFile
    } catch {
        Write-Error "Write_DetailedSystemInfo fehlgeschlagen: $_"
    }
}

function Finalize_LogSession {
    [CmdletBinding()]
    param(
        [string]$FinalizeMessage = "Log session completed normally"
    )

    if ($script:CurrentLogFile) {
        $separator = "=" * 80
        $footer = @(
            "",
            $separator,
            "LOG SESSION ENDED: $(Get-Date -Format $script:LogEntryFormat)",
            "Final Message: $FinalizeMessage",
            $separator
        ) -join "`r`n"

        Write_LogEntry -Message $FinalizeMessage -Level "SUCCESS"
        $footer | Out-File -FilePath $script:CurrentLogFile -Encoding UTF8 -Append

        # Session zurücksetzen
        $finalLogFile = $script:CurrentLogFile
        $script:CurrentLogFile = $null
        $script:SystemInfoFile = $null
        $script:CurrentProgram = $null
        $script:CurrentScriptType = $null
    } else {
        Write-Warning "Keine aktive Log-Sitzung vorhanden."
    }
}

function Get_LastLogFile {
    [CmdletBinding()]
    param(
        [string]$ProgramName
    )
    $p = if ($ProgramName) { Join-Path -Path $script:LogRootPath -ChildPath $ProgramName } else { $script:LogRootPath }
    if (-not (Test-Path $p)) { return $null }
    return Get-ChildItem -Path $p -File -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Get-GlobalErrorLog {
    [CmdletBinding()]
    param(
        [int]$LastNMonths = 1
    )

    $errorLogs = @()
    for ($i = 0; $i -lt $LastNMonths; $i++) {
        $date = (Get-Date).AddMonths(-$i)
        $errorLogPath = Join-Path -Path $script:LogRootPath -ChildPath "ERROR_$($date.ToString('yyyy-MM')).log"
        if (Test-Path $errorLogPath) {
            $errorLogs += Get-Item $errorLogPath
        }
    }

    return $errorLogs | Sort-Object LastWriteTime -Descending
}

Export-ModuleMember -Function *
