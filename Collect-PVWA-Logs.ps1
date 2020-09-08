###########################################################################
#
# SCRIPT NAME: Collect PVWA Logs
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.1 06/04/2107 - Update Logs and comments
# 1.2 09/08/2107 - Change From/To Type to string
# 1.3 14/08/2107 - Collecting more files
# 1.4 29/08/2107 - Minor bug fixes
# 1.5 18/10/2107 - Minor bug fixes
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk PVWA server logs

.DESCRIPTION
            This script can collect logs and configurations of a PVWA server
			
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

# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path

Function Get-WebConfigFromIIS
{
<# 
.SYNOPSIS 
	Method to return the web.config file of the PVWA

.DESCRIPTION
	This method will retrieve the PVWA web.config file path.
#>
	# Assuming running on Web Server
	# Get all Web Application Pools and filter only the Password Vault Web Access
	$iisPhysicalPath = $null
		
	if (Get-Module -ListAvailable -Name WebAdministration) {
		Write-LogMessage -MSG "Loading WebAdministration Module..." -Type Debug
		Import-Module WebAdministration -Verbose:$false
		Write-LogMessage -MSG "WebAdministration Module loaded" -Type Debug
	}
	else
	{ 
		Write-LogMessage -MSG "WebAdministration Module doesn't exists" -Type Debug 
		return $null
	}
	
	Foreach ($app in $(Get-WebApplication))
	{
		if($app.ApplicationPool -like "PasswordVault*") 
		{ 
			$iisPhysicalPath = $app.PhysicalPath;  
			break
		}
	}
	
	return $(Get-FilePath "$iisPhysicalPath\web.config")
}

Function Get-SettingFromIIS
{
<# 
.SYNOPSIS 
	Method to query settings from the PVWA IIS web.config file

.DESCRIPTION
	This method will collect settings from the PVWA web.config.
	Supported configuration names:
	 	VaultFile
	 	GWFile
	 	ConfigurationCredentialFile
		ConfigurationSafeName
	 	LogFolder		
.PARAMETER configToFetch
	The Setting name of configuration to retrieve
#>
	
	param ($configToFetch)
	
	$retConfigValue = $null
	$iisWebConfig = Get-WebConfigFromIIS
	if($configToFetch -ne "")
	{
		if($iisPhysicalPath -ne $null)
		{
			[xml]$WebConfig = Get-Content $iisWebConfig
			foreach ($key in $WebConfig.configuration.appSettings.add) 
			{ 
				if($key.key -eq $configToFetch) { $retConfigValue = $key.value }
			}
		}
	}
	else
	{
		$retConfigValue = $iisPhysicalPath
	}

	return $retConfigValue
}

Function Get-IISLogs
{

# Not implemented

}

Write-LogMessage -MSG "Collecting PVWA Files" -Type Debug
Write-LogMessage -MSG "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrPVWAFilePaths = @()

# Get the PVWA Installation Path
Write-LogMessage -MSG "Retrieving PVWA Log folder path from IIS" -Type Debug
$pvwaLogFolder = $(Get-SettingFromIIS -configToFetch "LogFolder")

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting PVWA file versions and additional information" -Type Debug
$pvwaVersions = "$DestFolderPath\_PVWAFileVersions.txt"
"PVWA: "+$(Get-FileVersion "$ComponentPath\Services\CyberArkScheduledTasks.exe") | Out-File $pvwaVersions
"Configuration Safe: "+$(Get-SettingFromIIS -configToFetch "ConfigurationSafeName") | Out-File $pvwaVersions -append
# Check that the logs folder is not empty
If([string]::IsNullOrEmpty($pvwaLogFolder))
{
	Write-LogMessage -MSG "PVWA Logs folder returned empty, assuming default 'C:\Windows\Temp\PVWA'" -Type Error
	$pvwaLogFolder = "C:\Windows\Temp\PVWA"
}	
"Log Folder: $pvwaLogFolder" | Out-File $pvwaVersions -append

Write-LogMessage -MSG "Collecting PVWA files list by timeframe" -Type Debug
$arrPVWAFilePaths += $pvwaVersions
If(Test-Path $pvwaLogFolder)
{
	$arrPVWAFilePaths += (Get-FilePath "$pvwaLogFolder\CyberArk.WebApplication.log")
	$arrPVWAFilePaths += (Get-FilePath "$pvwaLogFolder\CyberArk.WebConsole.log")
	$arrPVWAFilePaths += (Get-FilesByTimeframe -Path "$pvwaLogFolder\CyberArk.WebSession.*.log" -From $TimeframeFrom -To $TimeframeTo)
	$arrPVWAFilePaths += (Get-FilePath "$pvwaLogFolder\CyberArk.WebProfiling.log")
	$arrPVWAFilePaths += (Get-FilePath "$pvwaLogFolder\CyberArk.WebTasksEngine.log")
	$arrPVWAFilePaths += (Get-FilesByTimeframe -Path "$pvwaLogFolder\old\*.log" -From $TimeframeFrom -To $TimeframeTo)
}
$arrPVWAFilePaths += (Get-FilePath "$ComponentPath\Services\CyberArkScheduledTasks.exe.config")
$arrPVWAFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Services\Logs\*.log" -From $TimeframeFrom -To $TimeframeTo)
$arrPVWAFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Services\Logs\old\*.log" -From $TimeframeFrom -To $TimeframeTo)

$arrPVWAFilePaths += (Get-WebConfigFromIIS)
$arrPVWAFilePaths += (Get-IISLogs)

Collect-Files -arrFilesPath $arrPVWAFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting PVWA Files" -Type Debug