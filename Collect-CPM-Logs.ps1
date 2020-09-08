###########################################################################
#
# SCRIPT NAME: Collect CPM Logs
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.1 06/04/2107 - Update Logs and comments
# 1.2 09/08/2107 - Change From/To Type to string
# 1.3 29/08/2107 - Minor bug fixes
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk CPM server logs

.DESCRIPTION
            This script can collect logs and configurations of a CPM server
			
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

# Set DEP Registry path
$DEP_REG_PATH = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"

Function Get-DEPSettings
{
<# 
.SYNOPSIS 
	Method to query DEP settings

.DESCRIPTION
	This method will get the DEP settings from the computer including the exception files list
#>
	# DataExecutionPrevention_SupportPolicy  | Policy Level | Description
	# 2 | OptIn (default configuration) | Only Windows system components and services have DEP applied
	# 3 | OptOut | DEP is enabled for all processes. Administrators can manually create a list of specific applications which do not have DEP applied
	# 1 | AlwaysOn | DEP is enabled for all processes
	# 0 | AlwaysOff | DEP is not enabled for any processes
	
	$retDEPSettings = ""
	$depPolicyLevel = (gwmi Win32_OperatingSystem | Select DataExecutionPrevention_SupportPolicy).DataExecutionPrevention_SupportPolicy
	if ($depPolicyLevel -eq 3)
	{
		If (Test-Path $DEP_REG_PATH)
		{
			$depExclusions = Get-ItemProperty $DEP_REG_PATH | select * -ExcludeProperty PS*
			$depExclusionsList = $depExclusions.PSObject.Properties | Select Name
		}
	}
	Switch($depPolicyLevel)
	{
		0 { $retDEPSettings = "DEP is not enabled for any processes (Always Off)" }
		1 { $retDEPSettings = "DEP is enabled for all processes (Always On)" }
		2 { $retDEPSettings = "Only Windows system components and services have DEP applied (OptIn [default])" }
		3 { 
			$retDEPSettings = @()
			$retDEPSettings += "DEP is enabled for all processes. Administrators can manually create a list of specific applications which do not have DEP applied (OptOut)"
			$retDEPSettings += "Exclusions:"
			ForEach ($exc in $depExclusionsList)
			{
				$retDEPSettings += $exc.Name
			}
		  }
	}
	
	return $retDEPSettings
}

Write-LogMessage -MSG "Collecting CPM Files" -Type Debug
Write-LogMessage -MSG "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrCPMFilePaths = @()

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting CPM file versions and additional information" -Type Debug
$cpmVersions = "$DestFolderPath\_CPMFileVersions.txt"
"CPM:"+$(Get-FileVersion "$ComponentPath\PMEngine.exe") | Out-File $cpmVersions
"PMTerminal:"+$(Get-FileVersion "$ComponentPath\bin\PMTerminal.exe") | Out-File $cpmVersions -append
"PLink:"+$(Get-FileVersion "$ComponentPath\bin\plink.exe") | Out-File $cpmVersions -append
"DEP settings:"+$(Get-DEPSettings) | Out-File $cpmVersions -append

Write-LogMessage -MSG "Collecting CPM files list by timeframe" -Type Debug
$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Vault\Vault.ini")
$arrCPMFilePaths += $cpmVersions
$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Logs\*.log")
#$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Logs\pm_error.log")
#$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Logs\PMConsole.log")
#$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Logs\PMTrace.log")
#$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Logs\CACPMScanner.log")
$arrCPMFilePaths += (Get-FilePath "$ComponentPath\Scanner\Log\DNAConsole.log")
$arrCPMFilePaths += (Get-FilesByTimeframe -Path  "$ComponentPath\Scanner\Log\DNAtrace*.log" -From $TimeframeFrom -To $TimeframeTo)
$arrCPMFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Scanner\Log\MachineScans\*.log" -From $TimeframeFrom -To $TimeframeTo)
#$arrCPMFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\Archive\*.log" -From $TimeframeFrom -To $TimeframeTo)
$arrCPMFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\ThirdParty\*.log" -From $TimeframeFrom -To $TimeframeTo)
#$arrCPMFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\Old\ThirdParty\*.log" -From $TimeframeFrom -To $TimeframeTo)
#$arrCPMFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\Old\*.log" -From $TimeframeFrom -To $TimeframeTo)
$arrCPMFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\History\*.log" -From $TimeframeFrom -To $TimeframeTo)


Collect-Files -arrFilesPath $arrCPMFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting CPM Files" -Type Debug