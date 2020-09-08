###########################################################################
#
# SCRIPT NAME: Vault Disaster Recovery Logs
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk Vault server logs

.DESCRIPTION
            This script can collect logs and configurations of a Vault server
			
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
Function Get-InstallPath
{
<#
  .SYNOPSIS
  Get the installation path of a service
  .DESCRIPTION
  The function receive the service name from Detect-Component function and return the path or returns NULL if not found
  .EXAMPLE
  (Get-InstallPath $<ServiceName>) -ne $NULL
  .PARAMETER ServiceName
  The service name to query. Just one.
 #>

	param ($ServiceName)
	$retInstallPath = $Null
	if ($m_ServiceList -eq $null)
	{
		$m_ServiceList = Get-ChildItem "HKLM:\System\CurrentControlSet\Services" | ForEach-Object { Get-ItemProperty $_.pspath }
	}
	$regPath =  $m_ServiceList | Where-Object {$_.PSChildName -eq $ServiceName}
	If ($regPath -ne $Null)
	{
		$retInstallPath = $regPath.ImagePath.Substring($regPath.ImagePath.IndexOf('"'),$regPath.ImagePath.LastIndexOf('"')+1)
	}	
	return $retInstallPath
}

#endregion # Helper Functions

# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-LogMessage -MSG "Collecting DR Files" -Type Debug
Write-LogMessage -MSG "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrVaultFilePaths = @()

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting DR file versions and additional information" -Type Debug
$PADRVersions = "$DestFolderPath\_PADRFileVersions.txt"
"PADR:"+$(Get-FileVersion "$ComponentPath\PADR.exe") | Out-File $PADRVersions
$DRVersion = $(Get-Content $DestFolderPath\_PADRFileVersions.txt | Where-Object { $_.Contains("PADR:") }).Replace(".","")
$MyVerson = $DRVersion.Split(":")
$CompVer = $MyVerson[1].Substring(0,4)
Write-LogMessage -MSG "Collecting DR files list by timeframe" -Type Debug
$arrVaultFilePaths += $PADRVersions
$REGKEY_DR = "CyberArk Vault Disaster Recovery"
#
# get the path to the padr.ini of there and the vault is a HA Cluster Node
$DRINIPath = $(Get-InstallPath $REGKEY_DR).Replace($ComponentPath,"").Replace("\PADR.ini","").Replace("/service","").Replace("/DRFile","").Replace("PADR.exe","").Replace('"',"").Trim()
If ($CompVer -ge 1005) {
	$arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\padr.log")
	If ($DRINIPath -ne "") {
		#
		# v10.5 and newer HA Cluster Node
		$arrVaultFilePaths += (Get-FilePath "$DRINIPath\padr.ini")
	} else {
		#
		# v10.5 and newer Standalone DR
		$arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\padr.ini")
	}
} else {
	$arrVaultFilePaths += (Get-FilePath "$ComponentPath\padr.log")
	If ($DRINIPath -ne "") {
		#
		# v10.4 and older HA Cluster Node
		$arrVaultFilePaths += (Get-FilePath "$DRINIPath\padr.ini")
	} else {
		#
		# v10.4 and older Standalone DR
		$arrVaultFilePaths += (Get-FilePath "$ComponentPath\padr.ini")
	}
}

Collect-Files -arrFilesPath $arrVaultFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting Vault Files" -Type Debug