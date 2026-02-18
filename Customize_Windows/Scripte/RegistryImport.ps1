#https://stackoverflow.com/questions/49676660/how-to-run-the-reg-file-using-powershell
param (
    [Parameter(Mandatory = $true)]
    [string]$Path
)

function Test-IsAdmin {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Relaunch elevated once if not admin
if (-not (Test-IsAdmin)) {
    $cmd = Get-Command pwsh -ErrorAction SilentlyContinue
	if ($cmd) {
		$pwsh = $cmd.Source
	}
	else {
		$cmd2 = Get-Command powershell -ErrorAction SilentlyContinue
		if ($cmd2) { $pwsh = $cmd2.Source }
	}

	if (-not $pwsh) { throw "Cannot find pwsh or powershell to relaunch elevated." }

    $scriptPath = $MyInvocation.MyCommand.Definition
    $argList = @("-NoProfile","-ExecutionPolicy","Bypass","-File", "`"$scriptPath`"","-Path","`"$Path`"")
    Start-Process -FilePath $pwsh -ArgumentList $argList -Verb RunAs -WindowStyle Hidden
    Write-Host "Relaunching elevated..." -ForegroundColor Yellow
    exit
}

function RunRegistryImport {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Invalid path: $Path"
    }

    $regExe = Join-Path $Env:SystemRoot "System32\reg.exe"

    if (Test-Path $Path -PathType 'Leaf') {
        if ($Path -notlike '*.reg') { throw "File is not a .reg file: $Path" }
        ExecuteRegistryFile -FilePath $Path -RegExe $regExe
    }
    else {
        $regFiles = Get-ChildItem -LiteralPath $Path -Filter '*.reg' -File | Sort-Object Name
        if ($regFiles.Count -eq 0) {
            Write-Host "    No registry files found in the specified folder." -ForegroundColor Red
            return
        }
        foreach ($file in $regFiles) {
            ExecuteRegistryFile -FilePath $file.FullName -RegExe $regExe
        }
    }
}

function ExecuteRegistryFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $true)]
        [string]$RegExe
    )

    Write-Host "    Importing $([IO.Path]::GetFileName($FilePath)) ..." -NoNewline
    & $RegExe import $FilePath
    $code = $LASTEXITCODE

    if ($code -eq 0) {
        Write-Host " OK" -ForegroundColor Green
    } else {
        Write-Host " FAIL (exit $code)" -ForegroundColor Red
    }
}

RunRegistryImport -Path $Path