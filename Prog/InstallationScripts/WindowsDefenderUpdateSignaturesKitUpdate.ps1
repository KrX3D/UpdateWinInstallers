# DeployToolkit helpers
$dtPath = Join-Path $PSScriptRoot "Modules\DeployToolkit\DeployToolkit.psm1"
if (Test-Path $dtPath) {
    Import-Module -Name $dtPath -Force -ErrorAction Stop
} else {
    if (Get-Command -Name Write_LogEntry -ErrorAction SilentlyContinue) {
        Write_LogEntry -Message "DeployToolkit nicht gefunden: $dtPath" -Level "WARNING"
    } else {
        Write-Warning "DeployToolkit nicht gefunden: $dtPath"
    }
}

# Import shared configuration
$configPath = Join-Path -Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -ChildPath "Customize_Windows\Scripte\PowerShellVariables.ps1"

if (Test-Path -Path $configPath) {
    . $configPath # Import config file variables into current scope (shared server IP, paths, etc.)
} else {
    Write-Host ""
    Write-Host "Konfigurationsdatei nicht gefunden: $configPath" -ForegroundColor "Red"
    pause
    exit
}

# Dot-source the Functions.ps1 script
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
. "$NetworkShareDaten\Customize_Windows\Scripte\makecab.ps1"

$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

$webPageUrl = "https://www.microsoft.com/en-us/wdsi/defenderupdates"
$webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing
$content = $webPageContent.Content

# Function to extract version from the given HTML line using the specified pattern
function GetVersionFromLine($line, $pattern) {
    $match = [regex]::Match($line, $pattern)
    if ($match.Success) {
        return $match.Groups[1].Value
    }
    else {
        return "Not found"
    }
}

$lineRegex = '(?s)<li>Version: <span>(\d+\.\d+\.\d+\.\d+)</span><\/li>\s+<li>Engine Version: <span>(\d+\.\d+\.\d+\.\d+)</span><\/li>\s+<li>Platform Version: <span>(\d+\.\d+\.\d+\.\d+)</span><\/li>'
$lineMatch = [regex]::Match($content, $lineRegex)

if ($lineMatch.Success) {
    $intelligenceVersionWeb = $lineMatch.Groups[1].Value
    $engineVersionWeb = $lineMatch.Groups[2].Value
    $platformVersionWeb = $lineMatch.Groups[3].Value
}

$directory = "$NetworkShareDaten\Customize_Windows\Windows_Defender_Update_Iso\defender-update-kit-x64"

$cabFilePath = "$directory\defender-dism-x64.cab"
$xmlFileName = "package-defender.xml"
$extractPath = "$env:TEMP\defender"

$ProgramName = "Security intelligence Update Kit"

# Create the extraction directory if it doesn't exist
if (-not (Test-Path -Path $extractPath)) {
    New-Item -ItemType Directory -Path $extractPath | Out-Null
}

# Extract the CAB file using 7-Zip
& "C:\Program Files\7-Zip\7z.exe" x -o"$extractPath" "$cabFilePath" -y | Out-Null

# Path to the extracted XML file
$xmlFilePath = Join-Path $extractPath $xmlFileName

if (Test-Path $xmlFilePath) {
    # Read the XML file
    $xmlContent = Get-Content -Path $xmlFilePath -Raw

    # Load XML content
    $xml = [xml]$xmlContent

    # Extract the version information from XML
    $engineVersionCab = $xml.packageinfo.versions.engine
    $platformVersionCab = $xml.packageinfo.versions.platform
    $signaturesVersionCab = $xml.packageinfo.versions.signatures

    #Write-Host "CAB version information:"
    $cabFileVersions = @{
        "Platform version" = $platformVersionCab
        "Engine version" = $engineVersionCab
        "Security intelligence version" = $signaturesVersionCab
    }
    $cabFileVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Local"; Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize

    #Write-Host "Web page version information:"
    $webPageVersions = @{
        "Platform version" = $platformVersionWeb
        "Engine version" = $engineVersionWeb
        "Security intelligence version" = $intelligenceVersionWeb
    }
    $webPageVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Online"; Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize

    # Compare the versions
    $UpdateAvailable = $false
    Write-Host "Vergleich:" -foregroundcolor "DarkGray"
    foreach ($key in $webPageVersions.Keys) {
        $webVersion = $webPageVersions[$key]
        $cabVersion = $cabFileVersions[$key]
		
		if ([version]$webVersion -gt [version]$cabVersion) {
            #Write-Host "   $key is different. Web: $webVersion, CAB: $cabVersion" -foregroundcolor "Cyan"
            Write-Host "   $key is different:" -foregroundcolor "DarkGray"
            Write-Host "      CAB: $cabVersion" -foregroundcolor "Cyan"
            Write-Host "      Web: $webVersion" -foregroundcolor "Cyan"
            $UpdateAvailable = $true
        }
        else {
            Write-Host "   $key matches. Version: $webVersion" -foregroundcolor "DarkGray"
        }
    }

	Write-Host ""
	
    # Set $UpdateAvailable to true if any of the four elements is newer
    if ($UpdateAvailable) {
        Write-Host "Update ist vorhanden." -foregroundcolor "green"
		Write-Host ""
        # Extract the download link for the 64-bit version from the web page
		
		#$downloadLinkPattern = '<a class="c-hyperlink" href="(https:\/\/go\.microsoft\.com\/fwlink\/\?LinkID=121721&amp;arch=x64)">64-bit<\/a>'
        #$match = [regex]::Match($webPageContent.Content, $downloadLinkPattern)
        
        #if ($match.Success) {
            #$downloadLink = $match.Groups[1].Value
            $downloadLink = "https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64"
            #Write-Host "Download link for 64-bit version: $downloadLink"
			
            # Download the file to the temp folder
            $tempFilePath = Join-Path $env:TEMP "mpam-fe.exe"
            #Invoke-WebRequest -Uri $downloadLink -OutFile $tempFilePath
			$webClient = New-Object System.Net.WebClient
			[void](Invoke-DownloadFile -Url $downloadLink -OutFile $tempFilePath)
			$webClient.Dispose()

			#Write-Host "Extracting $tempFilePath using 7-Zip"
			$extractPathSub = Join-Path $extractPath "\Definition Updates\Updates"

			$arguments = "x `"$tempFilePath`" -o`"$extractPathSub`" -r -aoa"
			$process = New-Object System.Diagnostics.Process
			$process.StartInfo.FileName = $sevenZipPath
			$process.StartInfo.Arguments = $arguments
			$process.StartInfo.WindowStyle = 'Hidden'
			$process.StartInfo.UseShellExecute = $false
			$process.StartInfo.RedirectStandardOutput = $true
			$process.Start()
			$process.WaitForExit()

			# Capture the output
			$output = $process.StandardOutput.ReadToEnd()
			$process.Close()

			# Check if the extraction was successful
			if ($output -match "Everything is Ok") {
				#Write-Host "Extraction completed successfully."
				
				# Remove the MpSigStub.exe file
				$mpSigStubPath = Join-Path $extractPathSub "MpSigStub.exe"
				if (Test-Path $mpSigStubPath) {
					Remove-Item -Path $mpSigStubPath -Force
					#Write-Host "MpSigStub.exe file removed."
				}
				
				#Plattform aktualisieren				
				if($platformVersionCab -ne $platformVersionWeb)
				{
					#Write-Host "Eine neue Version von Plattform is vorhanden."
					#https://github.com/potatoqualitee/kbupdate

					# Install the 'kbupdate' module if not already present
					$moduleName = "kbupdate"
					if (-not (Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName })) {
						Install-Module -Name $moduleName -Force -Scope CurrentUser -AllowClobber
						#Write-Host "Module '$moduleName' wurde installiert." -foregroundcolor "Yellow"
					} else {
						#Write-Host "Module '$moduleName' is already installed."
					}

					# Import the 'kbupdate' module
					Import-Module $moduleName -ErrorAction Stop

					# Get the updates and find the newest version
					$updateList = Get-KbUpdate -Name KB4052623
					$newestUpdate = $updateList | Sort-Object -Property LastModified -Descending | Select-Object -First 1

					if ($newestUpdate) {
						# Find the link containing "updateplatform.amd64fre"
						$updateLink = $newestUpdate.Link | Where-Object { $_ -match "updateplatform\.amd64fre" }

						if ($updateLink) {
							# Define the download path to the desktop
							$tempPath = "$env:TEMP\"
							$downloadPath = Join-Path -Path $tempPath -ChildPath ($updateLink.Split("/")[-1])

							# Download the file to the desktop
							[void](Invoke-DownloadFile -Url $updateLink -OutFile $downloadPath)
						} else {
							Write-Host "Update link 'updateplatform.amd64fre' wurde nicht gefunden." -foregroundcolor "Red"
						}
					} else {
						Write-Host "Update wurde auf Windows Update Katalog nicht gefunden." -foregroundcolor "Red"
					}

					$updatePlatformFile = Get-ChildItem -Path $downloadPath -Filter "UpdatePlatform*.exe" -File -ErrorAction SilentlyContinue
					$updatePlatformFileVersion =  (Get-Item $updatePlatformFile).VersionInfo.ProductVersion

					if ($updatePlatformFile -and ($platformVersionCab -ne $updatePlatformFileVersion)) {
						#Write-Host "Entpacke $($updatePlatformFile.Name) " -foregroundcolor "Cyan"
						#Write-Host ""
						
						$path = $env:TEMP + "\defender\Platform"
						$currentFolder = (Get-ChildItem -Path $path -Directory | Select-Object -Last 1).Name
						$oldFolderPath = Join-Path -Path $path -ChildPath $currentFolder
						Remove-Item -Path $oldFolderPath -Force -Recurse
						
						$newFolderName = $platformVersionWeb + "-0"
						$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

						$arguments = "x `"$($updatePlatformFile.FullName)`" -o`"$path\$newFolderName`" -r"
						$process = New-Object System.Diagnostics.Process
						$process.StartInfo.FileName = $sevenZipPath
						$process.StartInfo.Arguments = $arguments
						$process.StartInfo.WindowStyle = 'Hidden'
						$process.StartInfo.UseShellExecute = $false
						$process.StartInfo.RedirectStandardOutput = $true
						$process.Start()
						$process.WaitForExit()

						# Capture the output
						$output = $process.StandardOutput.ReadToEnd()
						$process.Close()
						
						#Write-Host "Deleting $($updatePlatformFile.Name)"
						Remove-Item -Path $updatePlatformFile.FullName -Force
					}
				}
				
				#Edit the package-defender.xml file
				$packageXmlFile = "$extractPath\package-defender.xml"
				
				# Read the XML file
				$xml = [xml](Get-Content $packageXmlFile)

				# Update the values in the XML
				$xml.packageinfo.versions.signatures = $intelligenceVersionWeb
				$xml.packageinfo.versions.engine = $engineVersionWeb
				
				if ($platformVersionCab -ne $updatePlatformFileVersion)
				{
					$xml.packageinfo.versions.platform = $platformVersionWeb
				}

				# Save the updated XML back to the file
				$xml.Save($packageXmlFile)
				
			} else {
				#Write-Host "Extraction failed. Please check the 7-Zip command and arguments."
			}

			#Write-Host "Deleting mpam-fe.exe"
			Remove-Item -Path $tempFilePath -Force

			#Create cab file
			New-CAB -SourceDirectory $extractPath
			
			# Finde die zuletzt erstellte CAB im TEMP-Ordner
			$cabFile = Get-ChildItem -Path $env:TEMP -Filter "*.cab" -File |
					   Sort-Object LastWriteTime -Descending | Select-Object -First 1

			if (-not $cabFile) {
				Write-Host "Keine CAB-Datei im TEMP gefunden!"
				throw "CAB creation failed"
			}

			$cabFilePath = $cabFile.FullName
			Write-Host "Gefundene CAB: $cabFilePath"

			# einheitlicher Name
			$desiredName = Join-Path $env:TEMP "defender-dism-x64.cab"
			if ($cabFilePath -ne $desiredName) {
				try {
					Move-Item -Path $cabFilePath -Destination $desiredName -Force
					$cabFilePath = $desiredName
					Write-Host "CAB umbenannt/verschoben nach: $cabFilePath"
				} catch {
					Write-Host "Konnte CAB nicht umbenennen: $($_.Exception.Message)"
					# continue with original $cabFilePath
				}
			}

			# Delete DDF files
			$ddfFiles = Get-ChildItem -Path $env:TEMP -Filter "*.ddf" -File
			$ddfFiles | Remove-Item -Force
			
			# Pfad zum Signer-Skript (prüfen)
			$signerPath  = "$NetworkShareDaten\Customize_Windows\Scripte\certs\CreateCerts\Signer.ps1"
			if (-not (Test-Path $signerPath)) {
				Write-Host "Signer-Skript nicht gefunden: $signerPath" "ERROR"
				throw "Signer not found"
			}	

			# Explizite Windows PowerShell (64-bit) verwenden
			$winPs = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"

			# Optional: Unblock the signer in case it's blocked
			try { Unblock-File -Path $signerPath -ErrorAction SilentlyContinue } catch { }

			# --- Aufruf des Signer-Skripts (kein Start-Process, keine Files) ---
			Write-Host "Starte Signer: $signerPath mit Argument: $cabFilePath"
			
			# Fange stdout/stderr in-memory ab (2>&1) und führe synchron aus
			# Achtung: $output kann ein Array sein; wir joinden später mit "`n"
			$output = & $winPs -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $signerPath $cabFilePath 2>&1
			$exitCode = $LASTEXITCODE

			# Logge die Ausgabe
			if ($output) {
				# Wenn mehrere Zeilen, in einen String verwandeln
				$outText = if ($output -is [array]) { $output -join "`n" } else { [string]$output }
				Write-Host "Signer Ausgabe:`n$outText"
			}

			# Prüfe ExitCode
			if ($exitCode -eq 0) {
				Write-Host "Signer erfolgreich (ExitCode 0)"
			} else {
				Write-Host "Signer fehlgeschlagen (ExitCode $exitCode)"
				throw "Signer failed with exit code $exitCode"
			}

			# --- Move final CAB to destination ---
			try {
				Move-Item -Path $cabFilePath -Destination $directory -Force
				Write-Host "CAB nach $directory verschoben"
			} catch {
				Write-Host "Konnte CAB nicht verschieben: $($_.Exception.Message)"
				throw
			}

			Write-Host ""
			Write-Host "$ProgramName wurde aktualisiert.." -foregroundcolor "green"
        #}
        #else {
            #Write-Host "Download link not found."
        #}
    }
    else {
		Write-Host "Kein Update verfügbar. $ProgramName is aktuell." -foregroundcolor "DarkGray"
    }
}
else {
    #Write-Host "The XML file '$xmlFileName' was not found in the CAB archive."
}

# Clean up the extracted files
Remove-Item -Path $extractPath -Force -Recurse
# SIG # Begin signature block
# MIIc9gYJKoZIhvcNAQcCoIIc5zCCHOMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDK71/z94UF6GKe
# tqrC4Qt0uNgL1dbozMni6EcUq57V36CCFuwwggOuMIIClqADAgECAhBNocomnIW5
# s0j5bLoWINy9MA0GCSqGSIb3DQEBBQUAMG8xCzAJBgNVBAYTAkRFMQ8wDQYDVQQI
# DAZIZXNzZW4xETAPBgNVBAcMCFdlaWxidXJnMRIwEAYDVQQKDAlLclggQ29ycC4x
# CzAJBgNVBAsMAk9QMRswGQYDVQQDDBJXaW5kb3dzIFBvd2VyU2hlbGwwHhcNMjUw
# OTA0MTk0NDUzWhcNMjgwOTA0MTk0NDUzWjBvMQswCQYDVQQGEwJERTEPMA0GA1UE
# CAwGSGVzc2VuMREwDwYDVQQHDAhXZWlsYnVyZzESMBAGA1UECgwJS3JYIENvcnAu
# MQswCQYDVQQLDAJPUDEbMBkGA1UEAwwSV2luZG93cyBQb3dlclNoZWxsMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwQSdtHq5e48DRza31/FvoNG+O0k4
# sAxb2XZp5Y3Ab2MnouQ2+Wf4SWHMTKzyhvlF651mZrZUr8q0Jj0w5DXrEMmmsdBD
# dQ0ZFdn5N+8xdC8AQ+vmWjWS6WuUqEafFeJF9Vn8YvFzn7ybQQ6S01P6No0WCixu
# jzy18JNn0X7uil7WCkkHEoXx5zd7VLhGAvPp3gPrer8KoRTxOYH9Rtd08IBT3v6a
# pA1PfnY6AyLfSLo3609PTOMVn2tgqhvcfli+MeGw0cBY39tJPpl0fd4x+GGmXG+R
# 0oTTlwtI8mRZ5pgH6Gf4YOCmJm15Fpdn9MUHWJnq7TGMQF5OKdQz5ilCxQIDAQAB
# o0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFIYRyTxuV6F3qVuvVpAuK83/lgt4MA0GCSqGSIb3DQEBBQUAA4IBAQCZGPVq
# X9srByIWVSlUZyQoGW66CMAGdv9odr8IPwZ9aBtzAW7vq5HpkbL43/LjiZ7tvB74
# 24QYwdJziROVLDuS23Ms8Gi4VxFh3xsW1x7M/5cp4txcbXsC4iOSc0e0QvPBKx3s
# 8O6I//86xxf993S5X/WSJrHVjjaDt5RvRiyNgLRpb5HGCpqyo+SNAYZpL2XzEb9H
# 6lFIkmxPNSwCV7GEQesrdMNxIliKp46z29aEg2H0SDAlCiBigTz5vyggDnD1GUr/
# 6bQvqHKrL389gfo6t3ctGR8aXfEd9z7SEQ4oo3rircuWYG4G7Iy4l1zidZyd5w1K
# Gk79tpTVhhW3FTY2MIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkq
# hkiG9w0BAQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5
# WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJv
# b3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1K
# PDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2r
# snnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C
# 8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBf
# sXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGY
# QJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8
# rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaY
# dj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+
# wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw
# ++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+N
# P8m800ERElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7F
# wI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUw
# AwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAU
# Reuir/SSy4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEB
# BG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsG
# AQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAow
# CDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/
# Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLe
# JLxSA8hO0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE
# 1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9Hda
# XFSMb++hUD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbO
# byMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIG
# tDCCBJygAwIBAgIQDcesVwX/IZkuQEMiDDpJhjANBgkqhkiG9w0BAQsFADBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQw
# HhcNMjUwNTA3MDAwMDAwWhcNMzgwMTE0MjM1OTU5WjBpMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0
# ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMIICIjAN
# BgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAtHgx0wqYQXK+PEbAHKx126NGaHS0
# URedTa2NDZS1mZaDLFTtQ2oRjzUXMmxCqvkbsDpz4aH+qbxeLho8I6jY3xL1IusL
# opuW2qftJYJaDNs1+JH7Z+QdSKWM06qchUP+AbdJgMQB3h2DZ0Mal5kYp77jYMVQ
# XSZH++0trj6Ao+xh/AS7sQRuQL37QXbDhAktVJMQbzIBHYJBYgzWIjk8eDrYhXDE
# pKk7RdoX0M980EpLtlrNyHw0Xm+nt5pnYJU3Gmq6bNMI1I7Gb5IBZK4ivbVCiZv7
# PNBYqHEpNVWC2ZQ8BbfnFRQVESYOszFI2Wv82wnJRfN20VRS3hpLgIR4hjzL0hpo
# YGk81coWJ+KdPvMvaB0WkE/2qHxJ0ucS638ZxqU14lDnki7CcoKCz6eum5A19WZQ
# HkqUJfdkDjHkccpL6uoG8pbF0LJAQQZxst7VvwDDjAmSFTUms+wV/FbWBqi7fTJn
# jq3hj0XbQcd8hjj/q8d6ylgxCZSKi17yVp2NL+cnT6Toy+rN+nM8M7LnLqCrO2JP
# 3oW//1sfuZDKiDEb1AQ8es9Xr/u6bDTnYCTKIsDq1BtmXUqEG1NqzJKS4kOmxkYp
# 2WyODi7vQTCBZtVFJfVZ3j7OgWmnhFr4yUozZtqgPrHRVHhGNKlYzyjlroPxul+b
# gIspzOwbtmsgY1MCAwEAAaOCAV0wggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYD
# VR0OBBYEFO9vU0rp5AZ8esrikFb2L9RJ7MtOMB8GA1UdIwQYMBaAFOzX44LScV1k
# TN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcD
# CDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2lj
# ZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmww
# IAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUA
# A4ICAQAXzvsWgBz+Bz0RdnEwvb4LyLU0pn/N0IfFiBowf0/Dm1wGc/Do7oVMY2mh
# XZXjDNJQa8j00DNqhCT3t+s8G0iP5kvN2n7Jd2E4/iEIUBO41P5F448rSYJ59Ib6
# 1eoalhnd6ywFLerycvZTAz40y8S4F3/a+Z1jEMK/DMm/axFSgoR8n6c3nuZB9BfB
# wAQYK9FHaoq2e26MHvVY9gCDA/JYsq7pGdogP8HRtrYfctSLANEBfHU16r3J05qX
# 3kId+ZOczgj5kjatVB+NdADVZKON/gnZruMvNYY2o1f4MXRJDMdTSlOLh0HCn2cQ
# LwQCqjFbqrXuvTPSegOOzr4EWj7PtspIHBldNE2K9i697cvaiIo2p61Ed2p8xMJb
# 82Yosn0z4y25xUbI7GIN/TpVfHIqQ6Ku/qjTY6hc3hsXMrS+U0yy+GWqAXam4ToW
# d2UQ1KYT70kZjE4YtL8Pbzg0c1ugMZyZZd/BdHLiRu7hAWE6bTEm4XYRkA6Tl4KS
# FLFk43esaUeqGkH/wyW4N7OigizwJWeukcyIPbAvjSabnf7+Pu0VrFgoiovRDiyx
# 3zEdmcif/sYQsfch28bZeUz2rtY/9TCA6TD8dC3JE3rYkrhLULy7Dc90G6e8Blqm
# yIjlgp2+VqsS9/wQD7yFylIz0scmbKvFoW2jNrbM1pD2T7m3XDCCBu0wggTVoAMC
# AQICEAqA7xhLjfEFgtHEdqeVdGgwDQYJKoZIhvcNAQELBQAwaTELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBU
# cnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTAe
# Fw0yNTA2MDQwMDAwMDBaFw0zNjA5MDMyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcw
# FQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgU0hBMjU2
# IFJTQTQwOTYgVGltZXN0YW1wIFJlc3BvbmRlciAyMDI1IDEwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDQRqwtEsae0OquYFazK1e6b1H/hnAKAd/KN8wZ
# QjBjMqiZ3xTWcfsLwOvRxUwXcGx8AUjni6bz52fGTfr6PHRNv6T7zsf1Y/E3IU8k
# gNkeECqVQ+3bzWYesFtkepErvUSbf+EIYLkrLKd6qJnuzK8Vcn0DvbDMemQFoxQ2
# Dsw4vEjoT1FpS54dNApZfKY61HAldytxNM89PZXUP/5wWWURK+IfxiOg8W9lKMqz
# dIo7VA1R0V3Zp3DjjANwqAf4lEkTlCDQ0/fKJLKLkzGBTpx6EYevvOi7XOc4zyh1
# uSqgr6UnbksIcFJqLbkIXIPbcNmA98Oskkkrvt6lPAw/p4oDSRZreiwB7x9ykrjS
# 6GS3NR39iTTFS+ENTqW8m6THuOmHHjQNC3zbJ6nJ6SXiLSvw4Smz8U07hqF+8CTX
# aETkVWz0dVVZw7knh1WZXOLHgDvundrAtuvz0D3T+dYaNcwafsVCGZKUhQPL1naF
# KBy1p6llN3QgshRta6Eq4B40h5avMcpi54wm0i2ePZD5pPIssoszQyF4//3DoK2O
# 65Uck5Wggn8O2klETsJ7u8xEehGifgJYi+6I03UuT1j7FnrqVrOzaQoVJOeeStPe
# ldYRNMmSF3voIgMFtNGh86w3ISHNm0IaadCKCkUe2LnwJKa8TIlwCUNVwppwn4D3
# /Pt5pwIDAQABo4IBlTCCAZEwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQU5Dv88jHt
# /f3X85FxYxlQQ89hjOgwHwYDVR0jBBgwFoAU729TSunkBnx6yuKQVvYv1Ensy04w
# DgYDVR0PAQH/BAQDAgeAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMIGVBggrBgEF
# BQcBAQSBiDCBhTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MF0GCCsGAQUFBzAChlFodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcnQw
# XwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3Js
# MCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsF
# AAOCAgEAZSqt8RwnBLmuYEHs0QhEnmNAciH45PYiT9s1i6UKtW+FERp8FgXRGQ/Y
# AavXzWjZhY+hIfP2JkQ38U+wtJPBVBajYfrbIYG+Dui4I4PCvHpQuPqFgqp1PzC/
# ZRX4pvP/ciZmUnthfAEP1HShTrY+2DE5qjzvZs7JIIgt0GCFD9ktx0LxxtRQ7vll
# KluHWiKk6FxRPyUPxAAYH2Vy1lNM4kzekd8oEARzFAWgeW3az2xejEWLNN4eKGxD
# J8WDl/FQUSntbjZ80FU3i54tpx5F/0Kr15zW/mJAxZMVBrTE2oi0fcI8VMbtoRAm
# aaslNXdCG1+lqvP4FbrQ6IwSBXkZagHLhFU9HCrG/syTRLLhAezu/3Lr00GrJzPQ
# FnCEH1Y58678IgmfORBPC1JKkYaEt2OdDh4GmO0/5cHelAK2/gTlQJINqDr6Jfwy
# YHXSd+V08X1JUPvB4ILfJdmL+66Gp3CSBXG6IwXMZUXBhtCyIaehr0XkBoDIGMUG
# 1dUtwq1qmcwbdUfcSYCn+OwncVUXf53VJUNOaMWMts0VlRYxe5nK+At+DI96HAlX
# HAL5SlfYxJ7La54i71McVWRP66bW+yERNpbJCjyCYG2j+bdpxo/1Cy4uPcU3AWVP
# Grbn5PhDBf3Froguzzhk++ami+r3Qrx5bIbY3TVzgiFI7Gq3zWcxggVgMIIFXAIB
# ATCBgzBvMQswCQYDVQQGEwJERTEPMA0GA1UECAwGSGVzc2VuMREwDwYDVQQHDAhX
# ZWlsYnVyZzESMBAGA1UECgwJS3JYIENvcnAuMQswCQYDVQQLDAJPUDEbMBkGA1UE
# AwwSV2luZG93cyBQb3dlclNoZWxsAhBNocomnIW5s0j5bLoWINy9MA0GCWCGSAFl
# AwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJ
# KoZIhvcNAQkEMSIEIOa/Dz8/SarDQmz87yDQU2MprhnSC3J0K7EJhlF7GQu7MA0G
# CSqGSIb3DQEBAQUABIIBALUnhKFSr71CCTzYId4Ky1NgAgsTidgBXJS6+9Cri8q9
# 79Qbi+h3KAsElgXUHLVcIxSan45GI1sVEEMZQKvnr/vvps4NZT+tEgPZLAHX5VI7
# qboG519XCY0Kc1osnW4l+4MOYzP3Ghe2beWd+kFBJqbFLp5SnkFN7jR6p0SiU2wt
# Dk8O1HMehf/vqcrljQ+4ovHNzmjxHDuZ/QL4LM8tsDFDFEz31A9J+KCOXLU+R1Hj
# PSiY4doivpoEhCqt59cCQElBIKcicQ+oYP7aYHwwHYk1mQ9Cb+7z0Td6nx+6JXkl
# njKdNLDt+M3ZkWJUu7q2Th/dLEa9I36TbWB9YF8nWFuhggMmMIIDIgYJKoZIhvcN
# AQkGMYIDEzCCAw8CAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1w
# aW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExAhAKgO8YS43xBYLRxHanlXRoMA0G
# CWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjUwOTExMDkzODMzWjAvBgkqhkiG9w0BCQQxIgQgQR5CtFOSMHiR
# LMyge0bCev4EUjxm6pryUtzvmtzJe3UwDQYJKoZIhvcNAQEBBQAEggIAsjad/Ya6
# dHjJZm8PLg+I/p2ggDqTRLKXRSa0pFTon1q/fpSbTF0p3a6DwtYqjZH2DYlXJd8i
# NIOEnlvg/77EPA+TBGwG7X/eHJ2OBnOAOsVNM+BmMbN9VCn18bCDroGVF/RfBSY+
# oIm/qWm2+gFA4Zc5R3gaA9dRf0viqn1gIs6AzxwfnBBpcf8hPp9eXEoMq9nIxsRm
# J05yLTM4XW1w6cdBADxVHOV9BMvRwLZs4PLgrBVN7EzpOA8ZQZ0G0Gpe1bRC3ezq
# f2fMDSwd0QD/PWNlzs680G1Ibq2CVmYa9Lpz722+iuPabMhIjZbYG5x/+7eGN0VS
# vtyxX/nQLz37Bih0uyXGG+NaaZvrjVpx53KLyMmv9cMu1BV33Caws0mtMdILkrJN
# hUR1NoawfTwBc7bQp0UWWr62QW0Ca3qVI3vFDiBpfSDAEyVQ+CDyBXc1x3dmvclm
# i2funW544l11yERdZfpTgota/BkrB2m+H3L5x8gPAzGSwwHliLr8yBRzt8uGBMMQ
# SEuzVpssU7/QMt1iRNsZKuVxYxbw9IukOdy3TmLdjoyFCbkNR+gtOVA3eJaQf0P2
# buf1jMXBFVzNPDccx6UBR2OZIlvSft29BPmGMOcKu/ncbkw5qoqyKLPk7sxDTx9n
# a+fsTAKVoIo4mMwVs/GBgUKnkQSfNrSiSBo=
# SIG # End signature block
