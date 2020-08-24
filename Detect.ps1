###########################################################################
#
# SCRIPT NAME: Detect Components
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.5 28/08/2017 - Made more effecient, Added Log for the detect process
# 1.6 03/09/2017 - Support for detecting AIM
# 1.7 12/09/2017 - Support for detecting EPM Server
# 1.8 05/11/2017 - Fixed EPM Service Name
# 1.9 02/11/2020 - Support for detecting Cluster Vault Manager
# 1.10 26/02/2020 - Support for detecting RabbitMQ used in Distributed Vaults for PSM/PVWA
#
###########################################################################

<# 
.SYNOPSIS 
            A script to detect the CyberArk component server

.DESCRIPTION
            This script can detect the component server that this scripts runs on and returns the type of the component and the installation path
#>


# COMPONENTS SERVICE NAMES
$REGKEY_VAULTSERVICE_NEW = "CyberArk Logic Container"
$REGKEY_VAULTSERVICE_OLD = "Cyber-Ark Event Notification Engine"
$REGKEY_CPMSERVICE_NEW = "CyberArk Central Policy Manager Scanner"
$REGKEY_CPMSERVICE_OLD = "CyberArk Password Manager"
$REGKEY_PVWASERVICE = "CyberArk Scheduled Tasks"
$REGKEY_PSMSERVICE = "Cyber-Ark Privileged Session Manager"
$REGKEY_AIMSERVICE = "CyberArk Application Password Provider"
$REGKEY_EPMSERVICE = "VfBackgroundWorker"
$REGKEY_CVMSERVICE = "CyberArk Cluster Vault Manager"
$REGKEY_DRSERVICE = "CyberArk Vault Disaster Recovery"
$REGKEY_RMQSERVICE = "RabbitMQ"

# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path

# Set Log file path
$LOG_FILE_PATH = "$ScriptLocation\SupportTool_Detect.log"

# Set Debug / Verbose script modes
$InDebug = $PSBoundParameters.Debug.IsPresent
$InVerbose = $PSBoundParameters.Verbose.IsPresent

# Save the Services List
$m_ServiceList = $null

<#------------------------------------------------#>
<#----- Detect componenet installation path ------#>
<#------------------------------------------------#>

Function Log-MSG
{
<# 
.SYNOPSIS 
	Method to log a message on screen and in a log file

.DESCRIPTION
	Logging The input Message to the Screen and the Log File. 
	The Message Type is presented in colours on the screen based on the type

.PARAMETER MSG
	The message to log
.PARAMETER Header
	Adding a header line before the message
.PARAMETER SubHeader
	Adding a Sub header line before the message
.PARAMETER Footer
	Adding a footer line after the message
.PARAMETER Type
	The type of the message to log (Info, Warning, Error, Debug)
#>
	param(
		[Parameter(Mandatory=$true)]
		[AllowEmptyString()]
		[String]$MSG,
		[Parameter(Mandatory=$false)]
		[Switch]$Header,
		[Parameter(Mandatory=$false)]
		[Switch]$SubHeader,
		[Parameter(Mandatory=$false)]
		[Switch]$Footer,
		[Parameter(Mandatory=$false)]
		[ValidateSet("Info","Warning","Error","Debug","Verbose")]
		[String]$type = "Info"
	)
	
	If ($Header) {
		"=======================================" | Out-File -Append -FilePath $LOG_FILE_PATH 
		Write-Host "======================================="
	}
	ElseIf($SubHeader) { 
		"------------------------------------" | Out-File -Append -FilePath $LOG_FILE_PATH 
		Write-Host "------------------------------------"
	}
	
	$msgToWrite = "[$(Get-Date -Format "yyyy-MM-dd hh:mm:ss")]`t"
	$writeToFile = $true
	# Replace empty message with 'N/A'
	if([string]::IsNullOrEmpty($Msg)) { $Msg = "N/A" }
	# Check the message type
	switch ($type)
	{
		"Info" { 
			Write-Host $MSG.ToString()
			$msgToWrite += "[INFO]`t$Msg"
		}
		"Warning" {
			Write-Host $MSG.ToString() -ForegroundColor DarkYellow
			$msgToWrite += "[WARNING]`t$Msg"
		}
		"Error" {
			Write-Host $MSG.ToString() -ForegroundColor Red
			$msgToWrite += "[ERROR]`t$Msg"
		}
		"Debug" { 
			if($InDebug)
			{
				Write-Debug $MSG
				$msgToWrite += "[DEBUG]`t$Msg"
			}
			else { $writeToFile = $False }
		}
		"Verbose" { 
			if($InVerbose)
			{
				Write-Verbose $MSG
				$msgToWrite += "[VERBOSE]`t$Msg"
			}
			else { $writeToFile = $False }
		}
	}
	
	If($writeToFile) { $msgToWrite | Out-File -Append -FilePath $LOG_FILE_PATH }
	If ($Footer) { 
		"=======================================" | Out-File -Append -FilePath $LOG_FILE_PATH 
		Write-Host "======================================="
	}
}


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


Function Detect-Component
{
<#
  .SYNOPSIS
  Send the service name and recive service path
  .DESCRIPTION
  For each componet, the function send to Get-InstallPath function the service name and receive the path or NULL if not found
#>

	# Check if Vault is installed
	Log-MSG -Type "Debug" "Serching for Vault..."
	if(($(Get-InstallPath $REGKEY_VAULTSERVICE_OLD) -ne $NULL) -or ($(Get-InstallPath $REGKEY_VAULTSERVICE_NEW) -ne $NULL))
	{
		Log-MSG -Type "Info" "Found Vault installation"
		# Get the Vault Installation Path
		$vaultPath = $(Get-InstallPath $REGKEY_VAULTSERVICE_NEW).Replace("\LogicContainer\BLServiceApp.exe","").Replace('"',"").Trim()
		# For old versions
		# $vaultPath = $(Get-InstallPath $REGKEY_VAULTSERVICE_OLD).Replace("Event Notification Engine\ENE.exe","").Replace('"',"").Trim()
		Write-Host "Vault|$vaultPath"
	}
	# Check if CPM is installed
	Log-MSG -Type "Debug" "Serching for CPM..."
	if(($(Get-InstallPath $REGKEY_CPMSERVICE_OLD) -ne $NULL) -or ($(Get-InstallPath $REGKEY_CPMSERVICE_NEW) -ne $NULL))
	{
		# Get the CPM Installation Path
		Log-MSG -Type "Info" "Found CPM installation"
		$cpmPath = $(Get-InstallPath $REGKEY_CPMSERVICE_NEW).Replace("\Scanner\CACPMScanner.exe","").Replace('"',"").Trim()
		# For old versions
		# $cpmPath = $(Get-InstallPath $REGKEY_CPMSERVICE_OLD).Replace("PMEngine.exe").Replace("/SERVICE","").Replace('"',"").Trim()
		Write-Host "CPM|$cpmPath"
	}
	# Check if PVWA is installed
	Log-MSG -Type "Debug" "Serching for PVWA..."
	if($(Get-InstallPath $REGKEY_PVWASERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found PVWA installation"
		$pvwaPath = $(Get-InstallPath $REGKEY_PVWASERVICE).Replace("\Services\CyberArkScheduledTasks.exe","").Replace('"',"").Trim()
		Write-Host "PVWA|$pvwaPath"
	}
	# Check if PSM is installed
	Log-MSG -Type "Debug" "Serching for PSM..."
	if($(Get-InstallPath $REGKEY_PSMSERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found PSM installation"
		$PSMPath = $(Get-InstallPath $REGKEY_PSMSERVICE).Replace("CAPSM.exe","").Replace('"',"").Trim()
		Write-Host "PSM|$PSMPath"
	}
	# Check if AIM is installed
	Log-MSG -Type "Debug" "Serching for AIM..."
	if($(Get-InstallPath $REGKEY_AIMSERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found AIM installation"
		$AIMPath = $(Get-InstallPath $REGKEY_AIMSERVICE).Replace("/mode SERVICE","").Replace("\AppProvider.exe","").Replace('"',"").Trim()
		Write-Host "AIM|$AIMPath"
	}
	# Check if EPM Server is installed
	Log-MSG -Type "Debug" "Serching for EPM Server..."
	if($(Get-InstallPath $REGKEY_EPMSERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found EPM Server installation"
		$EPMPath = $(Get-InstallPath $REGKEY_EPMSERVICE).Replace("\VfBackgroundWorker.exe","").Replace('"',"").Trim()
		Write-Host "EPM|$EPMPath"
	}
	# Check if Cluster Vault Manager is installed
	Log-MSG -Type "Debug" "Serching for Cluster Vault Manager..."
	if($(Get-InstallPath $REGKEY_CVMSERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found Cluster Vault Manager installation"
		$CVMPath = $(Get-InstallPath $REGKEY_CVMSERVICE).Replace("\ClusterVault.exe","").Replace('"',"").Trim()
		Write-Host "CVM|$CVMPath"
	}
	# Check if Cluster Vault Manager is installed
	Log-MSG -Type "Debug" "Serching for Vault Disaster Recovery..."
	if($(Get-InstallPath $REGKEY_DRSERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found Vault Disaster Recover installation"
		$PADRPath = $(Get-InstallPath $REGKEY_DRSERVICE).Replace("/service","").Replace("/DRFile","").Replace("S:\PrivateArk\PADR\PADR.ini","").Replace("\PADR.exe","").Replace('"',"").Trim()
		Write-Host "PADR|$PADRPath"
	}
	# Check if RabbitMQ is installed Distributed Vaults for PSM/PVWA
	Log-MSG -Type "Debug" "Serching for RabbitMQ..."
	if($(Get-InstallPath $REGKEY_RMQSERVICE) -ne $NULL)
	{
		Log-MSG -Type "Info" "Found RabbitMQ installation"
		$RMQPath = $(Get-InstallPath $REGKEY_RMQSERVICE).Replace("\erlsrv.exe","").Replace('"',"").Trim()
		Write-Host "RabbitMQ|$RMQPath"
	}
}

Detect-Component