###########################################################################
#
# SCRIPT NAME: Collect Component Information
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.1 06/04/2107 - File collection fix, Write-LogMessage updates, Update Logs and comments
# 1.2 03/08/2107 - Adding more log messages, supporting Debug Switch for Debug messages, 
#				  Supporting ZIP files with no .Net requirement, Bug fixes
# 1.3 06/08/2017 - Fixing Zip method
# 1.4 09/08/2017 - Fixing Log collection by timeframe
# 1.5 14/08/2017 - Deleting the Temp folder and creating a LogOutput folder for the ZIP files
# 1.6 28/08/2017 - Fixing Log and adding Verbose Debug mode
# 1.7 30/08/2017 - Fixing minor bugs
# 1.8 03/09/2017 - Supporting AIM
# 1.9 04/09/2017 - Fixing minor bugs
# 2.0 12/09/2017 - Supporting EPM Server
# 2.1 18/09/2017 - Fixing minor bugs
# 2.2 24/10/2017 - Adding CYBRSupport Tool Version file to ZIP package
# 2.4 11/02/2020 - Supporting Cluster Vault Manager, Vault Disaster Recovery
# 2.5 26/02/2020 - New version - supporting RabbitMQ detection, configs, logs, and adding more vault logs/configs
# 2.6 09/04/2020 - add logic to the vault script to run diagnosedb report when database service is running
# 2.7 28/05/2020 - collecting event log information specific to the PSM - System, Security, Application, AppLocker
#
###########################################################################
<# 
.SYNOPSIS 
      The main script to collect CyberArk component logs and configurations

.DESCRIPTION
      This script will initialize commands and scripts to collect relevant information on a component server.
			Supported components are: Vault, CPM, PVWA, PSM, AIM, EPM Server, Cluster Vault Manager
#>

[CmdletBinding(DefaultParametersetName='Collect')] 
param
(
	[Parameter(Mandatory=$true,HelpMessage="Enter the Component Name")]
	[ValidateSet("Vault","CPM","PVWA","PSM","AIM","EPM","CVM","PADR","RabbitMQ")]
	[Alias("server")]
	[String]$ComponentName,
	
	# Use this switch to collect component information
	[Parameter(ParameterSetName='Collect',Mandatory=$false)][switch]$Collect,
	[Parameter(ParameterSetName='Collect',Mandatory=$true,HelpMessage="Enter the installation path of the component")]
	[String]$InstallPath,
	
	[Parameter(ParameterSetName='Collect',Mandatory=$false,HelpMessage="Collect OS Environment information")]
	[Switch]$OSInfo,
	
	[Parameter(ParameterSetName='Collect',Mandatory=$false,HelpMessage="Collect Event Viewer information")]
	[Switch]$EventViewer,
	
	[Parameter(ParameterSetName='Collect',Mandatory=$false,HelpMessage="Timeframe collection From date")]
	#[ValidateScript({(Get-Date $_) -le (Get-Date)})]
	[Alias("From")]
	[String]$TimeframeFrom,
	[Parameter(ParameterSetName='Collect',Mandatory=$false,HelpMessage="Timeframe collection To date")]
	#[ValidateScript({(Get-Date $_) -ge (Get-Date)})]
	[Alias("To")]
	[String]$TimeframeTo,
	
	[Parameter(ParameterSetName='Collect',Mandatory=$false,HelpMessage="Enter any additional paths to collect")]
	[String[]]$AdditionalPaths,
	
	# Use this switch to analyse component information
	[Parameter(ParameterSetName='Analyse',Mandatory=$false)][switch]$Analyse
	
)

# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path
# Set Zip Location
$ZipLocationFolder = "$ScriptLocation\LogsOutput"
# Get Debug / Verbose parameters for Script
$InDebug = $PSBoundParameters.Debug.IsPresent
$InVerbose = $PSBoundParameters.Verbose.IsPresent
# Script Version
$ScriptVersion = "2.8"

# Set Log file path
$global:LOG_FILE_PATH = "$ScriptLocation\SupportTool.log"
# Set Date Time Pattern
[string]$global:g_DateTimePattern = "$([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.ShortDatePattern) $([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.LongTimePattern)"
# Set Module Path
$MODULE_SHARED = "$ScriptLocation\bin\CYBRSupportTool-Shared.psd1"
# Set Scripts Paths
$VAULT_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-VAULT-Logs.ps1"
$CPM_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-CPM-Logs.ps1"
$PVWA_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-PVWA-Logs.ps1"
$PSM_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-PSM-Logs.ps1"
$AIM_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-AIM-Logs.ps1"
$EPM_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-EPM-Logs.ps1"
$CVM_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-CVM-Logs.ps1"
$DR_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-PADR-Logs.ps1"
$RMQ_LOG_COLLECTION_PATH = "$ScriptLocation\Collect-RabbitMQ-Logs.ps1"

#region Helper Functions
# @FUNCTION@ ======================================================================================================================
# Name...........: Load-Modules
# Description....: Load the relevant modules into the script
# Parameters.....: None
# Return Values..: None
# =================================================================================================================================
Function Load-Modules
{
<# 
.SYNOPSIS 
	Load Support Tool modules
.DESCRIPTION
	Load all relevant Support Tool modules for the script
#>
	param(
	)

	$shared = Import-Module $MODULE_SHARED -Force -DisableNameChecking -PassThru -ErrorAction Stop 

	return $shared
}

# @FUNCTION@ ======================================================================================================================
# Name...........: UnLoad-Modules
# Description....: UnLoad the relevant modules into the script
# Parameters.....: Module Info
# Return Values..: None
# =================================================================================================================================
Function UnLoad-Modules
{
<# 
.SYNOPSIS 
	UnLoad hardening modules
.DESCRIPTION
	UnLoad all relevant hardening modules for the script
#>
	param(
		$moduleInfo
	)

	ForEach ($info in $moduleInfo)
	{
		Remove-Module -ModuleInfo $info -ErrorAction Stop | out-Null
	}
}


Function Zip-Files
{	
<# 
.SYNOPSIS 
	Method to ZIP a folder to a ZIP file

.DESCRIPTION
	Method to ZIP all files from a folder to a destination ZIP file

.PARAMETER zipFileName
	The ZIP file path
.PARAMETER sourceDir
	The folder path from which files will be collected to a ZIP
#>
	param($zipFileName, $sourceDir)
	
	# Add the CYBRSupportTool_Version to the directory before the ZIP
	Copy-Item -Path "$ScriptLocation\CYBRSupportTool_Version.txt" -Destination $sourceDir
	
	try
	{
		# Works only with .Net Framework 4 and above
		[Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" ) | out-null
		[System.IO.Compression.ZipFile]::CreateFromDirectory($sourceDir, $zipFileName)
	} catch {
		# Falling back to old school ZIP
		If(-not (Test-Path($zipFileName)))
		{
			set-content $zipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
			#get the zip file object
			$zipfile = $zipFileName | Get-Item -ErrorAction Stop
	 
			#make sure it is not set to ReadOnly
			Write-Verbose "Setting isReadOnly to False"
			$zipfile.IsReadOnly = $false 	
		}
		
		$shellApplication = new-object -com shell.application
		$zipPackage = $shellApplication.NameSpace($zipfile.FullName)
		$filesToCopy = (Get-ChildItem $sourceDir | Sort Length)
		foreach($file in $filesToCopy) 
		{ 
			$zipPackage.CopyHere($file.FullName)
			Start-sleep -milliseconds ($file.length/1000 + 150)
		}
	}
}
#endregion


#region Specific Components Scripts
Function Collect-ComponentInfo
{
<#
 .SYNOPSIS
 Collect Component Specific logs and configuration
 .DESCRIPTION
 Collect Component Specific logs and configuration to an output folder
 .PARAMETER Name
 The Name of the Component to collect date from
 According to this name, the method will run the relevant collection script
 .PARAMETER Path
 The Path of installation of the required component
 .PARAMETER OutFolder
 The path of the output folder to save the logs and configurations
 .PARAMETER LogsFrom
 Retrive log entries from a specific date
 .PARAMETER LogsTo
 Retrive log entries to a specific date
#>
	param(
		[Parameter(Mandatory=$true)]
		[String]$Name,
		[Parameter(Mandatory=$true)]
		[String]$Path,
		[Parameter(Mandatory=$true,HelpMessage="Enter the output folder path")]
		[ValidateScript({Test-Path $_})]
		[String]$OutFolder,
		[Parameter(Mandatory=$true)]
		[Alias("From")]
		[String]$LogsFrom,
		[Parameter(Mandatory=$true)]
		[Alias("To")]
		[String]$LogsTo
	)
	
	Write-LogMessage -MSG "Start running component script" -SubHeader -LogFile $LOG_FILE_PATH
	If ([string]::IsNullOrEmpty($LogsFrom)) { $LogsFrom = $(Get-Date (Get-Date).AddHours(-72) -Format $g_DateTimePattern) } 
	If ([string]::IsNullOrEmpty($LogsTo)) { $LogsTo = $(Get-Date -Format $g_DateTimePattern) }
	Write-LogMessage -MSG "Will collect logs from $LogsFrom to $LogsTo" -Type Debug -LogFile $LOG_FILE_PATH
	$scriptPathAndArgs = "& `"{0}`" -ComponentPath '$Path' -DestFolderPath '$OutFolder' -TimeframeFrom `"$LogsFrom`" -TimeframeTo `"$LogsTo`""
	$scriptExpression = ""
		
	Switch($Name)
	{
		"Vault" { $scriptExpression = ($scriptPathAndArgs -f $VAULT_LOG_COLLECTION_PATH) }
		"CPM" { $scriptExpression = ($scriptPathAndArgs -f $CPM_LOG_COLLECTION_PATH) }
		"PVWA" { $scriptExpression = ($scriptPathAndArgs -f $PVWA_LOG_COLLECTION_PATH) }
		"PSM" { 
			$scriptExpression = ($scriptPathAndArgs -f $PSM_LOG_COLLECTION_PATH)
       
			Write-LogMessage -MSG "Exporting PSM Security Event Viewer logs from $TimeframeFrom to $TimeframeTo" -Type Debug -LogFile $LOG_FILE_PATH
			Get-EventViewer -Log "Security" -OutFolder $tempFolder -From $TimeframeFrom -To $TimeframeTo
			Write-LogMessage -MSG "Exporting PSM AppLocker Event Viewer logs from $TimeframeFrom to $TimeframeTo" -Type Debug -LogFile $LOG_FILE_PATH
			. $ScriptLocation\Get-AppLockerEvent.ps1
			$AppLockerFile = "$OutFolder\AppLocker-EventLog.csv"
			Get-AppLockerEvent -StartDate $TimeFrameTo.AddDays(-3) -Full | Export-Csv $AppLockerFile
		}
		"AIM" { $scriptExpression = ($scriptPathAndArgs -f $AIM_LOG_COLLECTION_PATH) }
		"EPM" { $scriptExpression = ($scriptPathAndArgs -f $EPM_LOG_COLLECTION_PATH) }
		"CVM" { $scriptExpression = ($scriptPathAndArgs -f $CVM_LOG_COLLECTION_PATH) }
		"DR" { $scriptExpression = ($scriptPathAndArgs -f $DR_LOG_COLLECTION_PATH) }
		"RabbitMQ" { $scriptExpression = ($scriptPathAndArgs -f $RMQ_LOG_COLLECTION_PATH) }
	}
	Write-LogMessage -MSG "Running $Name Collection script" -Type Debug -LogFile $LOG_FILE_PATH
	Write-LogMessage -MSG $scriptExpression -Type Debug -LogFile $LOG_FILE_PATH
	Invoke-Expression $scriptExpression
	Write-LogMessage -MSG "Done running script" -SubHeader -LogFile $LOG_FILE_PATH
}
#endregion

#---------------
# Check if Powershell is running in Constrained Language Mode
If($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage")
{
	Write-LogMessage -Type Error -Msg "Powershell is currently running in $($ExecutionContext.SessionState.LanguageMode) mode which limits the use of some API methods used in this script.`
	PowerShell Constrained Language mode was designed to work with system-wide application control solutions such as CyberArk EPM or Device Guard User Mode Code Integrity (UMCI).`
	For more information: https://blogs.msdn.microsoft.com/powershell/2017/11/02/powershell-constrained-language-mode/"
	Write-LogMessage -Type Info -Msg "Script ended"
	return
}

# Load all relevant modules
$moduleInfos = Load-Modules

Write-LogMessage -Type Info -MSG "Starting script (v$ScriptVersion)" -Header -LogFile $LOG_FILE_PATH
if($InDebug) { Write-LogMessage -Type Info -MSG "Running in Debug Mode" -LogFile $LOG_FILE_PATH }
if($InVerbose) { Write-LogMessage -Type Info -MSG "Running in Verbose Mode" -LogFile $LOG_FILE_PATH }
Write-LogMessage -Type Debug -MSG "Running PowerShell version $($PSVersionTable.PSVersion.Major) compatible of versions $($PSVersionTable.PSCompatibleVersions -join ", ")" -LogFile $LOG_FILE_PATH
# Verify the Powershell version is compatible
If (!($PSVersionTable.PSCompatibleVersions -join ", ") -like "*3*")
{
	Write-LogMessage -Type Error -Msg "The Powershell version installed on this machine is not compatible with the required version for this script.`
	Installed PowerShell version $($PSVersionTable.PSVersion.Major) is compatible with versions $($PSVersionTable.PSCompatibleVersions -join ", ").`
	Please install at least PowerShell version 3."
	Write-LogMessage -Type Info -Msg "Script ended"
	return
}

If($Analyze)
{
	Throw "This is not yet supported"
	exit
}

If($Collect)
{
	# Write the CYBRSupport Tool Version file
	"CYBRSupportTool Running version $ScriptVersion" | Out-File -FilePath "$ScriptLocation\CYBRSupportTool_Version.txt" -Force
	Write-LogMessage -Type Info -MSG "Collecting $ComponentName relevant information" -LogFile $LOG_FILE_PATH
	
	$zipFileName = "$ComponentName-$(Get-Date -Format 'ddMMyyyy_hhmmss')"
	$tempFolder = New-Item -Type "Directory" -Path "$($ENV:TEMP)\$zipFileName"
		
	If([string]::IsNullOrEmpty($TimeframeFrom)) { $TimeframeFrom = $(Get-Date (Get-Date).AddHours(-72) -Format $g_DateTimePattern) }
	If([string]::IsNullOrEmpty($TimeframeTo)) { $TimeframeTo = $(Get-Date -Format $g_DateTimePattern) }

	# Get OS Info
	If($OSInfo) { Get-OSInfo -OutFolder $tempFolder }
	# Get Event Viewer
	If($EventViewer) 
	{ 
		Write-LogMessage -MSG "Exporting Application Event Viewer logs from $TimeframeFrom to $TimeframeTo" -LogFile $LOG_FILE_PATH
		Get-EventViewer -Log "Application" -OutFolder $tempFolder -From $TimeframeFrom -To $TimeframeTo
		Write-LogMessage -MSG "Exporting SystemEvent Viewer logs from $TimeframeFrom to $TimeframeTo" -LogFile $LOG_FILE_PATH
		Get-EventViewer -Log "System" -OutFolder $tempFolder -From $TimeframeFrom -To $TimeframeTo
	}
	
	# Check that the installation path exists
	If(Test-Path $InstallPath)
	{
		Collect-ComponentInfo -Name $ComponentName -Path $InstallPath -OutFolder $tempFolder -LogsFrom $TimeframeFrom -LogsTo $TimeframeTo
		If($AdditionalPaths.Count -gt 0)
		{
			Foreach($path in $AdditionalPaths)
			{
				If(Test-Path $path)
				{
					Collect-ComponentInfo -Name $ComponentName -Path $path -OutFolder $tempFolder -LogsFrom $TimeframeFrom -LogsTo $TimeframeTo
				}
				Else
				{ Throw "$ComponentName additional path ($path) does not exit" }
			}
		}
		# Check if we have the output logs folder
		If(!(Test-Path $ZipLocationFolder))
		{
			# Folder does not exists - Create it
			New-Item -Type "Directory" -Path "$ZipLocationFolder"
		}
		Write-LogMessage -MSG "Zipping files from $tempFolder to $ZipLocationFolder\$zipFileName.zip" -LogFile $LOG_FILE_PATH
		Zip-Files -zipFileName "$ZipLocationFolder\$zipFileName.zip" -sourceDir $tempFolder
		Write-LogMessage -MSG "Zip file size $(Format-Number (Get-Item "$ZipLocationFolder\$zipFileName.zip").Length)" -LogFile $LOG_FILE_PATH
		# Delete the Temp folder
		Write-LogMessage -MSG "Deleting temp folder $tempFolder" -LogFile $LOG_FILE_PATH
		Remove-Item -Recurse -Force -Path $tempFolder
	}
	Else
	{
		Throw "$ComponentName installation path ($InstallPath) does not exit"
		Exit
	}
}

Write-LogMessage -Type Info -MSG "Script ended" -Footer -LogFile $LOG_FILE_PATH

# UnLoad loaded modules
UnLoad-Modules $moduleInfos