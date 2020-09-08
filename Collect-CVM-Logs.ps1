###########################################################################
#
# SCRIPT NAME: Collect Vault Logs
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk Cluster Vault Manager logs

.DESCRIPTION
            This script can collect logs and configurations of a Cluster Vault Manager
			
.PARAMETER ComponentPath
The installation path of the Component
.PARAMETER DestFolderPath
Folder to export the Logs (Path)
.PARAMETER TimeframeFrom
Time frame (From) 
.PARAMETER TimeframeTo
Time frame (To) 
#>

param
(
	[Parameter(Mandatory=$true,HelpMessage="Enter the Component Name")]
	[ValidateScript({Test-Path $_})]
	[String]$ComponentPath,
	
	[Parameter(Mandatory=$true,HelpMessage="Enter the Destination folder path to save the logs")]
	[ValidateScript({Test-Path $_})]
	[String]$DestFolderPath,
	
	[Parameter(Mandatory=$false,HelpMessage="Timeframe collection From date")]
	#[ValidateScript({(Get-Date $_) -le (Get-Date)})]
	[Alias("From")]
	[String]$TimeframeFrom,
	[Parameter(Mandatory=$false,HelpMessage="Timeframe collection To date")]
	#[ValidateScript({(Get-Date $_) -ge (Get-Date)})]
	[Alias("To")]
	[String]$TimeframeTo
)

#region Helper Functions
#endregion # Helper Functions

# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-LogMessage -MSG "Collecting Cluster Vault Manager Files" -Type Debug
Write-LogMessage -MSG "Collecting Cluster Vault Manager logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrVaultFilePaths = @()

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting Cluster Vault Manager file versions and additional information" -Type Debug
$vaultVersions = "$DestFolderPath\_ClusterVaultManagerFileVersions.txt"
"ClusterVault:"+$(Get-FileVersion "$ComponentPath\ClusterVault.exe") | Out-File $vaultVersions
$CVMVersion = Get-Content $DestFolderPath\_ClusterVaultManagerFileVersions.txt | Where-Object { $_.Contains("ClusterVault:") }
Write-LogMessage -MSG "$CVMVersion" -Type Info
$MyVerson = $CVMVersion.Split(":")
$CompVer = $MyVerson[1].Replace(".","").Substring(0,4)
Write-LogMessage -MSG "Collecting Cluster Vault Manager files list by timeframe" -Type Debug
$arrVaultFilePaths += $vaultVersions
If ($CompVer -ge 1005) {
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\ClusterVault.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\ClusterVaultDynamics.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\ClusterVaultConsole.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\ClusterVaultTrace.log")
} else {
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\ClusterVault.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\ClusterVaultDynamics.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\ClusterVaultConsole.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\ClusterVaultTrace.log")
}

Collect-Files -arrFilesPath $arrVaultFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting Cluster Vault Manager Files" -Type Debug