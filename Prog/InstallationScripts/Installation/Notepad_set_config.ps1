# DeployToolkit helpers
$dtPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Modules\DeployToolkit\DeployToolkit.psm1"
if (Test-Path $dtPath) {
    Import-Module -Name $dtPath -Force -ErrorAction Stop
} else {
    if (Get-Command -Name Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "DeployToolkit nicht gefunden: $dtPath" -Level "WARNING"
    } else {
        Write-Warning "DeployToolkit nicht gefunden: $dtPath"
    }
}

﻿# Path to your config.xml
$configPath = "$env:APPDATA\Notepad++\config.xml"

# 1) Launch & graceful close so config.xml is generated
$p = Start-Process notepad++ -PassThru
Start-Sleep 3
$p.CloseMainWindow() | Out-Null
if (-not $p.WaitForExit(10000)) { $p.Kill(); $p.WaitForExit() }

# 2) Verify file exists
if (-not (Test-Path $configPath)) {
    Write-Error "config.xml not found at $configPath. Please start & close Notepad++ manually once."
    exit 1
}

# 3) Read file as lines
$lines = Get-Content $configPath

# flags to know what we’ve seen
$foundToolBar = $false
$foundNewDoc  = $false
$foundBackup  = $false
$foundDark    = $false

# 4) Process each line
$newLines = foreach ($line in $lines) {
    $out = $line

    # a) ToolBar
    if ($line -match '<GUIConfig\s+name="ToolBar"') {
        $foundToolBar = $true
        #Write-Host "Found ToolBar line:`n  $line"
        # change fluentCustomColor and inner text
        $out = $out -replace 'fluentCustomColor="\d+"','fluentCustomColor="16229180"'
        $out = $out -replace '(>)[^<]*(</GUIConfig>)', '$1small$2'
        #Write-Host "Changed to:     `n  $out`n"
    }

    # b) NewDocDefaultSettings
    if ($line -match '<GUIConfig\s+name="NewDocDefaultSettings"') {
        $foundNewDoc = $true
        #Write-Host "Found NewDocDefaultSettings line:`n  $line"
        $out = $out -replace 'encoding="[^"]*"','encoding="1"'
        $out = $out -replace 'openAnsiAsUTF8="[^"]*"','openAnsiAsUTF8="no"'
        #Write-Host "Changed to:               `n  $out`n"
    }

    # c) Backup
    if ($line -match '<GUIConfig\s+name="Backup"') {
        $foundBackup = $true
        #Write-Host "Found Backup line:`n  $line"
        $out = $out -replace 'action="[^"]*"','action="0"'
        # ensure self‑closing
        $out = $out -replace '>\s*</GUIConfig>', ' />'
        #Write-Host "Changed to:        `n  $out`n"
    }

    # d) DarkMode
    if ($line -match '<GUIConfig\s+name="DarkMode"') {
        $foundDark = $true
        #Write-Host "Found DarkMode line:`n  $line"
        $out = $out -replace 'enable="[^"]*"','enable="yes"'
        #Write-Host "Changed to:       `n  $out`n"
    }

    $out
}

# 5) If any were missing, inject them under <GUIConfigs>
if (-not $foundToolBar) {
    Write-Warning "ToolBar entry not found; injecting one."
    $injection = '    <GUIConfig name="ToolBar" visible="yes" fluentColor="0" fluentCustomColor="16229180" fluentMono="no">small</GUIConfig>'
    $newLines = $newLines -replace '(<GUIConfigs>.*)', "`$1`r`n$injection"
}
if (-not $foundNewDoc) {
    Write-Warning "NewDocDefaultSettings entry not found; injecting one."
    $injection = '    <GUIConfig name="NewDocDefaultSettings" format="0" encoding="1" lang="0" codepage="-1" openAnsiAsUTF8="no" addNewDocumentOnStartup="no" />'
    $newLines = $newLines -replace '(<GUIConfigs>.*)', "`$1`r`n$injection"
}
if (-not $foundBackup) {
    Write-Warning "Backup entry not found; injecting one."
    $injection = '    <GUIConfig name="Backup" action="0" useCustumDir="no" dir="" isSnapshotMode="yes" snapshotBackupTiming="7000" />'
    $newLines = $newLines -replace '(<GUIConfigs>.*)', "`$1`r`n$injection"
}
if (-not $foundDark) {
    Write-Warning "DarkMode entry not found; injecting one."
    $injection = '    <GUIConfig name="DarkMode" enable="yes" />'
    $newLines = $newLines -replace '(<GUIConfigs>.*)', "`$1`r`n$injection"
}

# 6) Write back without adding or removing any other lines
$newLines | Set-Content $configPath

#Write-Host "All done. Four GUIConfig entries updated (or injected)."
