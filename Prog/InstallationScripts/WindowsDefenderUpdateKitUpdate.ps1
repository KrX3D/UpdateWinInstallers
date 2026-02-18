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

$webPageUrl = "https://support.microsoft.com/en-us/topic/microsoft-defender-update-for-windows-operating-system-installation-images-1c89630b-61ff-00a1-04e2-2d1f3865450d"
$webPageContent = Invoke-WebRequest -Uri $webPageUrl -UseBasicParsing

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

# Extract version information from the HTML lines
$defenderPackageLine = $webPageContent.Content | Select-String -Pattern "Defender package version:"
$platformVersionLine = $webPageContent.Content | Select-String -Pattern "Platform version:"
$engineVersionLine = $webPageContent.Content | Select-String -Pattern "Engine version:"
$intelligenceVersionLine = $webPageContent.Content | Select-String -Pattern "Security intelligence version:"

# Updated regex patterns
#$defenderPackageVersionPattern = '>\s*(\d+(?:\.\d+){1,3})\s*<'
$defenderPackageVersionPattern = 'Defender package version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
#$platformVersionPattern = '<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
$platformVersionPattern = 'Platform version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
$engineVersionPattern = 'Engine version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
#$engineVersionPattern = 'Engine version: <b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
$intelligenceVersionPattern = 'Security intelligence version:\s*<b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'
#$intelligenceVersionPattern = 'Security intelligence version: <b class="ocpLegacyBold">(\d+\.\d+\.\d+\.\d+)</b>'

$defenderPackageVersionWeb = GetVersionFromLine $defenderPackageLine $defenderPackageVersionPattern
$platformVersionWeb = GetVersionFromLine $platformVersionLine $platformVersionPattern
$engineVersionWeb = GetVersionFromLine $engineVersionLine $engineVersionPattern
$intelligenceVersionWeb = GetVersionFromLine $intelligenceVersionLine $intelligenceVersionPattern

$directory = "$NetworkShareDaten\Customize_Windows\Windows_Defender_Update_Iso\defender-update-kit-x64"

$cabFilePath = "$directory\defender-dism-x64.cab"
$xmlFileName = "package-defender.xml"
$extractPath = "$env:TEMP\defender"

$ProgramName = "Windows Defender Update Kit"

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
    $defenderVersionCab = $xml.packageinfo.versions.defender
    $engineVersionCab = $xml.packageinfo.versions.engine
    $platformVersionCab = $xml.packageinfo.versions.platform
    $signaturesVersionCab = $xml.packageinfo.versions.signatures

    #Write-Host "CAB version information:"
    $cabFileVersions = @{
        "Defender package version" = $defenderVersionCab
        "Platform version" = $platformVersionCab
        "Engine version" = $engineVersionCab
        "Security intelligence version" = $signaturesVersionCab
    }
    $cabFileVersions.GetEnumerator() | Sort-Object Name | Format-Table -Property @{Label = "Local"; Expression = {$_.Key}}, @{Label = "Version"; Expression = {$_.Value}} -AutoSize

    #Write-Host "Web page version information:"
    $webPageVersions = @{
        "Defender package version" = $defenderPackageVersionWeb
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
        #$downloadLinkPattern = '<a href="(https://go\.microsoft\.com/fwlink/\?linkid=2144531)" target="_blank" class="ocpExternalLink">64-bit</a>'
        #$match = [regex]::Match($webPageContent.Content, $downloadLinkPattern)
		
		$downloadLink = "https://go.microsoft.com/fwlink/?linkid=2144531"
        
        #if ($match.Success) {
            #$downloadLink = $match.Groups[1].Value
            #Write-Host "Download link for 64-bit version: $downloadLink"
			
            # Download the file to the temp folder
            $tempFilePath = Join-Path $env:TEMP "defender-update-kit-x64.cab"
            #Invoke-WebRequest -Uri $downloadLink -OutFile $tempFilePath
			$webClient = New-Object System.Net.WebClient
			$webClient.DownloadFile($downloadLink, $tempFilePath)
			$webClient.Dispose()

            # Remove old files from the defender-update-kit-x64 folder
            Remove-Item -Path "$directory\*" -Force -Recurse

            # Extract the downloaded file using 7-Zip
            & "C:\Program Files\7-Zip\7z.exe" x -o"$directory" $tempFilePath -y | Out-Null

            # Remove the downloaded file
            Remove-Item -Path $tempFilePath -Force
			
			#Edit downloaded PS1 file to skip signature check
			$filePath = "$directory\DefenderUpdateWinImage.ps1"
			
			# Read the file content
			$fileContent = Get-Content -Path $filePath

			# Search for the line to comment out
			$lineToComment = '    ValidateCodeSign -PackageFile $PkgFile'
			$commentedLine = '#' + $lineToComment
			
			# Loop through each line and find the line to comment out
			for ($i = 0; $i -lt $fileContent.Length; $i++) {
				if ($fileContent[$i] -eq $lineToComment) {
					# Comment out the line
					$fileContent[$i] = $commentedLine

					# Exit the loop since we found and modified the line
					break
				}
			}
			
			# Write the modified content back to the file
			$fileContent | Set-Content -Path $filePath

			#Write-Host "The line has been commented out in the file."

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDubo0WaGsW5/qN
# CN+GTfIlXbwBcDfV/IMDbRQPKdxEjKCCFuwwggOuMIIClqADAgECAhBNocomnIW5
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
# KoZIhvcNAQkEMSIEIIznLzaOC62G2kLunzYjpQaQmfhSKQ0ckoaSkRxYhxqEMA0G
# CSqGSIb3DQEBAQUABIIBAIstfnlEm1cZVgcXBXH9T1CQuJzglPxrBYMwTi5NEeZI
# 88eenz/cdZY87UV3d3cjin4PlhvYxO4wQP+6GxswzTkWfyKaU5SWliSisPzgIEv3
# Or3sjhhjI36ye8/oo5s4FFaYiU0IQTDtOfekCYJSPyJIwMqylYVUJzKlTZsxexfy
# BqGd4jXdPpCueFWenPyPGWlAb0fdiq3pTsN0wxumwKvg9RTIM/UQYcV7/QtulIkc
# +TuOI1JNoeBbkPh7LGAWy3ObA+Z6UrghLXY9KIlRm0HnVZ7aU1XYk/MYkXi/TyW/
# uebN2VwBLm+hXpMbFEpD4wYZkNuwtqLB4iqLszRVOFOhggMmMIIDIgYJKoZIhvcN
# AQkGMYIDEzCCAw8CAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1w
# aW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExAhAKgO8YS43xBYLRxHanlXRoMA0G
# CWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMjUwOTExMDkzNzI1WjAvBgkqhkiG9w0BCQQxIgQgOJEG7XiOW7Er
# TaVBEpnCpf/1MUjeD3Gb9okfNfzH12cwDQYJKoZIhvcNAQEBBQAEggIAn+ZQjejW
# R1edHK3hH+YeOHyOW0eR9ZLsVwHr5N0clp8TFxFLmLGtB5iwNEz3PB/DfNseyfAy
# 6Rj4MpneeLn5YiCALNM3QSgZxhmm6OCJXK1RUnvAIn8iPpoeZDSKa/qwRYo4N/IV
# mnbE4ogqQ1r8ieGkfX9doyn7e4CAjLK9vvx38Xj4b/2jmnDHvARnrg9eQ8N2bCGT
# IojmR7noZkYaCAiRpRkf3oZCk+/8NjEQReGY/bl+DwDq21DVn9vIX4J3X50II2SI
# mItK6D3iaavq4hDU2Ok3QZKZfBZUnrxqMd4r446DCeR3/dTCUG22+ARwj8Kr5+9G
# 1CSVCxVw9CfAvCa9Mayyz1N9eM3Iyk7yp0sq26RjQBOZ0d+naqTZ1GfmY1eo5uKa
# dXSDUSjmlCmczLfq8jQFg/pAHQIj9dB1gEeb/EFsFx64EExBBrOU4/05tJznjKVg
# ExfdKdGRMB+71+wEKZxZrzbcxAGCJzdoeDaT6yYHFvA8AQI+GSQZNMo3ucknu+Pv
# OxSUDVccnaCsj4Zz7zY9Gi4lkxnYBSkejegnxbhwm9/FIFG+rkTn0otppoP42KmT
# u1hLSziaiiFOzUdji5y+W8MHslYaEdNKMUYwGfJXxxvbQ1cUx69fe/mZAe3ApnZ4
# dHPV47EzU5bcS38OmLKFVrY2ZOlc1FHPDGw=
# SIG # End signature block
