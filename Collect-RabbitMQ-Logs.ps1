###########################################################################
#
# SCRIPT NAME: RabbitMQ logs and configurations
#
# VERSION HISTORY:
# 1.0 26/02/2020 - Initial release
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the RabbitMQ files

.DESCRIPTION
            This script can collect logs and configurations of RabbitMQ when used for Distributed Vaults and PSM/PVWA
			
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

Write-LogMessage -MSG "Collecting RabbitMQ Files" -Type Debug

$arrVaultFilePaths = @()
#
# get the base path for RabbitMQ to collect the information
$RabbitMQPath=$ComponentPath.Replace("\erl10.2\erts-10.2\bin","").Trim()

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting RabbitMQ file versions and additional information" -Type Debug
$vaultVersions = "$DestFolderPath\_RabbitMQFileVersions.txt"
"RabbitMQ:"+$(Get-FileVersion "$ComponentPath\erlsrv.exe") | Out-File $vaultVersions
"RabbitMQConfiguratgor:"+$(Get-FileVersion "$RabbitMQPath\RabbitMQ Configurator\RabbitMQConfigurator.exe") | Out-File $vaultVersions

#
# paths to logs and configuration files
$arrVaultFilePaths += (Get-FilePath "$ComponentPath\erl.ini")
$arrVaultFilePaths += (Get-FilePath "$RabbitMQPath\logs\*.log")
$arrVaultFilePaths += (Get-FilePath "$RabbitMQPath\RabbitMQ Configurator\RabbitMQConfigurator.exe.config")
$arrVaultFilePaths += (Get-FilePath "$RabbitMQPath\RabbitMQ Configurator\Conf\*.*")

Collect-Files -arrFilesPath $arrVaultFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting Cluster Vault Manager Files" -Type Debug