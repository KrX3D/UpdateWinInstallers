$ip = "192.168.1.10"
$Serverip = "\\$ip"
$NetworkShareDaten = "$Serverip\Files"
$InstallationFolder = "$NetworkShareDaten\Prog"

#Git Token für api Abfrage um max Abfragen zu erhöhen
$GitHubToken = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

#Get the path for Powershell 7 and if not installed Powershell 5
$helperUNC = "$NetworkShareDaten\Customize_Windows\Scripte\GetPowerShellExe.ps1"
$scriptText = Get-Content -Path $helperUNC -Raw
$sb = [ScriptBlock]::Create($scriptText)
$PSHostPath = & $sb