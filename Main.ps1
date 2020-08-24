###########################################################################
#
# SCRIPT NAME: Collect Component Information
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.1 06/04/2107 - File collection fix, Log-Msg updates, Update Logs and comments
# 1.2 03/08/2107 - Adding more log messages, supporting Debug Switch for Debug messages, 
#				   Supporting ZIP files with no .Net requirement, Bug fixes
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
$ScriptVersion = "2.7"

# Set Log file path
$LOG_FILE_PATH = "$ScriptLocation\SupportTool.log"
# Set Date Time Pattern
[string]$g_DateTimePattern = "$([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.ShortDatePattern) $([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.LongTimePattern)"
# Set Registry Keys
$REG_COMPUTER_NAME = "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" # \ComputerName
$REG_PROCESSOR_NAME = "HARDWARE\DESCRIPTION\System\CentralProcessor\0" # \ProcessorNameString
$REG_HKLM_ENIVORNMENT = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment" # \NUM_OF_PROCESSOR, \PROCESSOR_ARCHITECTURE
$REG_USER_NAME = "Software\Microsoft\Windows\CurrentVersion\Explorer" # HKCU \Logon User Name
$REG_USER_DOMAIN = "SOFTWARE\Policies\Microsoft\System\DNSClient" # PrimaryDnsSuffix
$REG_HKCU_ENIVORNMENT = "Volatile Environment" # \USERDNSDOMAIN

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

#region Helper functions
Function Log-MSG
{
<# 
.SYNOPSIS 
	Method to log a message on screen and in a log file

.DESCRIPTION
	Logging The input Message to the Screen and the Log File. 
	The Message Type is presented in colours on the screen based on the type

.PARAMETER LogFile
	The Log File to write to. By default using the LOG_FILE_PATH
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
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
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
		[String]$type = "Info",
		[Parameter(Mandatory=$false)]
		[String]$LogFile = $LOG_FILE_PATH
	)
	
	If ($Header) {
		"=======================================" | Out-File -Append -FilePath $LogFile 
		Write-Host "======================================="
	}
	ElseIf($SubHeader) { 
		"------------------------------------" | Out-File -Append -FilePath $LogFile 
		Write-Host "------------------------------------"
	}
	
	$msgToWrite = "[$(Get-Date -Format "yyyy-MM-dd hh:mm:ss")]`t"
	$writeToFile = $true
	# Replace empty message with 'N/A'
	#if([string]::IsNullOrEmpty($Msg)) { $Msg = "N/A" }
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
			if($InDebug -or $InVerbose)
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
	
	If($writeToFile) { $msgToWrite | Out-File -Append -FilePath $LogFile }
	If ($Footer) { 
		"=======================================" | Out-File -Append -FilePath $LogFile 
		Write-Host "======================================="
	}
}

Function Check-Service
{
<# 
.SYNOPSIS 
	Method to query a service status

.DESCRIPTION
	Returns the Status of a Service, Using Get-Service.
.PARAMETER ServiceName
	The Service Name to Check Status for
#>
	param (
		$ServiceName
	)
	$ErrorActionPreference = "SilentlyContinue"
	$svcStatus = "" # Init
	# Create command to run
	$svcStatus = Get-Service -Name $ServiceName | Select Status
	$ErrorActionPreference = "Continue"
	Log-MSG -Type "Debug" -Msg "$ServiceName Service Status is: $($svcStatus.Status)"  -LogFile $LOG_FILE_PATH
	return $svcStatus.Status
}

Function Check-RegistryService($strComputer)
{
<# 
.SYNOPSIS 
	Method to query the RemoteRegistry service

.DESCRIPTION
	Returns the Status of The Remote Registry Service, Using Get-Service.
#>
	Return Check-Service RemoteRegistry
}

Function Get-Reg 
{
<# 
.SYNOPSIS 
	Method that will connect to a remote computer Registry using the Parameters it receives

.DESCRIPTION
	Returns the Value Data of the Registry Value Name Queried on a remote machine

.PARAMETER Hive
	The Hive Name (LocalMachine, Users, CurrentUser)
.PARAMETER Key
	The Registry Key Path
.PARAMETER Value
	The Registry Value Name to Query
.PARAMETER RemoteComputer
	The Computer Name that we want to Query (Default Value is Local Computer)
#>
	param($Hive,
		$Key,
		$Value,
		$RemoteComputer="." # If not entered Local Computer is Selected
	)
	# Catch Exceptions
	$ErrorActionPreference="SilentlyContinue"
	trap [Exception] { if($verbose){ LOG-MSG -Msg "$RemoteComputer Has Registry Problems:`n`t$_" -Type "Error"   -LogFile $LOG_FILE_PATH }; continue; }	
	# Connect to Remote Computer Registry
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $RemoteComputer)
	if($reg -eq $null) { return "Registry Access Error"; } # Registry Access Error - Exit Function
	# Open Remote Sub Key
	$regKey= $reg.OpenSubKey($Key)
	if($Value -eq $null) # Enumerate Keys
	{ return $regKey.GetSubKeyNames() } # Return Sub Key Names
	if($regKey.ValueCount -gt 0) # check if there are Values 
	{ return $regKey.GetValue($Value) } # Return Value
}

Function Get-WMIItem {
<# 
.SYNOPSIS 
	Method Retrieves a specific Item from a remote computer's WMI

.DESCRIPTION
	Returns the Value Data of a specific WMI query on a remote machine

.PARAMETER Class
	The WMI Class Name
.PARAMETER Item
	The Item to query
.PARAMETER Query
	A WMI query to run
.PARAMETER Filter
	A filter item to filter the results
.PARAMETER RemoteComputer
	The Computer Name that we want to Query (Default Value is Local Computer)
#>
	param($Class,
		$RemoteComputer=".", # If not entered Local Computer is Selected
		$Item,
		$Query="", # If not entered an empty WMI SQL Query is Entered
		$Filter="" # If not entered an empty Filter is Entered
	)
	$ErrorActionPreference="SilentlyContinue"
	trap [Exception] { return "WMI Error";continue }
	if ($Query -eq "") # No Specific WMI SQL Query
	{
		# Execute WMI Query, Return only the Requested Items
		gwmi -Class $Class -ComputerName $RemoteComputer -Filter $Filter -Property $Item | Select $Item
	}
	else # User Entered a WMI SQL Query
	{gwmi -ComputerName $RemoteComputer -Query $Query | select $Item}
	$ErrorActionPreference="Continue"
}

Function Get-FileVersion
{
<# 
.SYNOPSIS 
	Method to return a file version

.DESCRIPTION
	Returns the File version and Build number
	Returns Null if not found

.PARAMETER filePath
	The path to the file to query
#>
	param ($filePath)
	$retFileVersion = $Null
	$path = resolve-path $filePath 
	If ($path -ne $null)
	{
		$retFileVersion = ($path | Get-Item | select VersionInfo).VersionInfo.ProductVersion
	}
	
	return $retFileVersion
}

Function Get-FilePath
{
<# 
.SYNOPSIS 
	Method to return a file (resolved) path

.DESCRIPTION
	Returns the File path (if exists)

.PARAMETER filePath
	The path to the file to query
#>
	param($filePath)
	
	try
	{
		If(![String]::IsNullOrEmpty($filePath) -and (Test-Path $filePath))
		{
			Log-MSG -Type Debug -Msg "Retrieving file: $filePath"  -LogFile $LOG_FILE_PATH
			return (Resolve-Path $filePath)
		}
		else
		{
			Log-Msg -Type Error -Msg "File $filePath not found"  -LogFile $LOG_FILE_PATH
		}
	}
	catch
	{
		Log-Msg -Type Error -Msg $_  -LogFile $LOG_FILE_PATH
	}
}

Function Get-DateFromParameter
{
<# 
.SYNOPSIS 
	Method to return a DateTime object from a string Date Time

.DESCRIPTION
	a DateTime object from a string Date Time (on any format and in any Culture)

.PARAMETER sDate
	A String Date Time
#>
	param ([string]$sDate)
	$retDate = $null
	$retDate = [System.DateTime]::ParseExact($sDate,$g_DateTimePattern,[System.Globalization.CultureInfo]::InvariantCulture)

	return $retDate
}

Function Get-FilesByTimeframe
{
<# 
.SYNOPSIS 
	Method to return a file (resolved) path accourding to a time frame

.DESCRIPTION
	Returns the File path (if exists) only if the last update time of the file falls between the specified time frame
	This method adds the timezone of the machine to the time frame so that the time frame will cover for time differences as well

.PARAMETER path
	The path to the file/folder to query
.PARAMETER from
	A date from which to filter the files
.PARAMETER to
	A date to which to filter the files
#>
	param ($path, $from, $to)
	$arrFilePaths = @()
	try
	{
		$BiasHours = [System.TimezoneInfo]::Local.BaseUtcOffset.Hours
		$FromTimezone = $(Get-DateFromParameter $From).AddHours(-$BiasHours)
		$ToTimezone = $(Get-DateFromParameter $To).AddHours($BiasHours)
		If(![string]::IsNullOrEmpty($path) -and (Test-Path $path))
		{
			$selectedFiles = Get-ChildItem -Path $path | Where-Object {(! $_.PSIsContainer) -and ($_.LastWriteTime -gt $FromTimezone) -and ($_.LastWriteTime -le $ToTimezone)} | Select FullName
			ForEach ($file in $selectedFiles)
			{
				Log-MSG "Adding File $($file.FullName) to Files collection" -Type Debug  -LogFile $LOG_FILE_PATH
				$arrFilePaths += (Get-FilePath $file.FullName)
			}
			return $arrFilePaths
		}
		else
		{
			Log-Msg -Type Error -Msg "The Path received ($path) to collect files by time frame (From: $from; To: $to;) was empty or does not exist"  -LogFile $LOG_FILE_PATH
		}
	}
	catch
	{
		Log-Msg -Type Error -Msg $_	 -LogFile $LOG_FILE_PATH										   
	}
}

Function Collect-Files
{
<# 
.SYNOPSIS 
	Method to collect all files from a list

.DESCRIPTION
	Method collects all files from a list to a destination folder

.PARAMETER arrFilesPath
	A list of file paths to collect
.PARAMETER destFolder
	The target folder to which files will be collected
#>
	
	param ($arrFilesPath, $destFolder)
		
	Log-Msg -Msg "Collecting Files to $destFolder..."  -LogFile $LOG_FILE_PATH
	try {
		# Copy all files to the temp folder
		ForEach ($filePath in $arrFilesPath)
		{
			If(![String]::IsNullOrEmpty($filePath) -and (Test-Path $filePath))
			{															 
				If($filePath.GetType().Name -eq "String")
				{
					if ($filePath.IndexOf("*") -gt 0)
					{ 
						If(Test-Path $filePath)
						{
							dir $filePath | ?{!$_.PsIsContainer} | sort | %{ Copy -Path $_.FullName -Destination $destFolder }
						}
					}
				}
				else
				{ Copy -Path $filePath -Destination $destFolder }
			}
			else
			{
				Log-Msg -Type Error -Msg "The Path received ($filePath) to collect files was empty or does not exist"  -LogFile $LOG_FILE_PATH
			}
		}
		
		Log-Msg "Completed collecting files!"  -LogFile $LOG_FILE_PATH
	}
	catch
	{
		Log-Msg -Type Error -Msg $_
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

Function Format-Number($num)
{
<# 
.SYNOPSIS 
	Method to Format a size number to KB/MB/GB

.DESCRIPTION
	Method to Format a size number to KB/MB/GB

.PARAMETER Num
	An unfromatted number
#>
	if($num -ge 1GB)
		{ return ($num/1GB).ToString("# GB") } # Format to GB
	elseif($num -ge 1MB)
		{ return ($num/1MB).ToString("# MB") } # Format to MB
	else
		{ return ($num/1KB).ToString("# KB") } # Format to KB
}
#endregion

#region Collect OS Environment information
Function Get-ComputerDetails
{
<# 
.SYNOPSIS 
	Method to Collect machine infomration

.DESCRIPTION
	Method to collect machine information using WMI or registry.
	Some of the infomration is available only vai WMI and not Registry
	The Information collected:
		Machine Name
		Domain Name
		System type
		Manufacturer name (according to BIOS information)
		Model (according to BIOS information)
		Number of Processors
		Phisical Memory (Total and Availble)
		User Name

.PARAMETER (Switch) scanWMI
	Whether to scan using WMI (True / False)
	(Preffered method)
.PARAMETER (Switch) scanReg
	Whether to scan using Registry (True / False)
#>
	param([switch]$scanWMI, [switch]$scanReg)
	# Create an Object With Empty Values
	$ComputerDet = "" | Select Caption,Domain,SystemType,Manufacturer,Model,NumberOfProcessors,TotalPhysicalMemory,AvialableMem,UserName
	If($scanWMI)
	{
		# Collect Computer Details from Win32_computersystem Using WMI
		$wmiDet = Get-WMIItem -Class "Win32_computersystem" -Item Caption,Domain,SystemType,Manufacturer,Model,NumberOfProcessors,TotalPhysicalMemory,UserName
		$ComputerDet.Caption = $wmiDet.Caption
		$ComputerDet.Domain = $wmiDet.Domain
		$ComputerDet.SystemType = $wmiDet.SystemType
		$ComputerDet.Model = $wmiDet.Model
		$ComputerDet.NumberOfProcessors = $wmiDet.NumberOfProcessors
		$ComputerDet.TotalPhysicalMemory = $wmiDet.TotalPhysicalMemory
		$ComputerDet.UserName = $wmiDet.UserName
		$ComputerDet.Manufacturer = $wmiDet.Manufacturer.Replace(","," ")
		# Check Total Physical Memory Size and Format it accordingly
		$ComputerDet.TotalPhysicalMemory = Format-Number $wmiDet.TotalPhysicalMemory
		# Collect Available Memory with WMI
		$wmiAvialableMem = Get-WMIItem -Class "Win32_PerfFormattedData_PerfOS_Memory" -Item "AvailableBytes"
		$ComputerDet.AvialableMem = Format-Number $wmiAvialableMem.AvailableBytes
	}
	ElseIf($scanReg)
	{
		# Get The Needed Information From Registry
		if($verbose) { Log-MSG "`tSome of the Computers Details will not be collected" }
		$ComputerDet.Caption = Get-Reg -Hive LocalMachine -Key $REG_COMPUTER_NAME -Value "ComputerName"
		if($ComputerDet.Caption -eq "Registry Access Error") 
		{ 
			$ComputerDet.Caption = $Env:ComputerName
			$ComputerDet.Domain = "N/A"
		}
		else{
			$ComputerDet.NumberOfProcessors = Get-Reg -Hive LocalMachine -Key $REG_HKLM_ENIVORNMENT -Value "NUMBER_OF_PROCESSORS"
			$ComputerDet.UserName = Get-Reg -Hive CurrentUser -Key $REG_USER_NAME -Value "Logon User Name"
			$ComputerDet.Domain = Get-Reg -Hive CurrentUser -Key $REG_USER_DOMAIN -Value "PrimaryDnsSuffix"
			if($ComputerDet.Domain -ne $Null -and $ComputerDet.UserName -ne $Null)
			{
				$ComputerDet.UserName = $ComputerDet.Domain.split(".")[0]+"\"+$ComputerDet.UserNam
			}
			$ComputerDet.SystemType = Get-Reg -Hive LocalMachine -Key $REG_HKLM_ENIVORNMENT -Value "PROCESSOR_ARCHITECTURE"
			if($ComputerDet.SystemType -like "x86") { $ComputerDet.SystemType = "X86-based PC" }
			elseif($ComputerDet.SystemType -like "x64") { $ComputerDet.SystemType = "x64-based PC" }
			else { $ComputerDet.SystemType = "Unknown" }
		}		
	}
	
	return $ComputerDet
}

Function Get-CPUDetails
{
<# 
.SYNOPSIS 
	Method to Collect processor infomration

.DESCRIPTION
	Method to collect processor information using WMI or registry.
	The Information collected:
		Processor Name(s)
		
.PARAMETER (Switch) scanWMI
	Whether to scan using WMI (True / False)
	(Preffered method)
.PARAMETER (Switch) scanReg
	Whether to scan using Registry (True / False)
#>
	param([switch]$scanWMI, [switch]$scanReg)
	if($scanWMI) # Scan WMI
	{
		# Collect CPU Name Using WMI
		$CPUName = Get-WMIItem -Class "Win32_Processor" -Item Name
		# CPU Names Can Contain Multiple Values
		$arrCPUNames = @() 
		foreach($CPU in $CPUName){
			$arrCPUNames += $CPU.Name.Trim() # the String of the CPU Name has White Space in The Beginning - Trim It
		}
	}
	elseif($scanReg) # Scan Registry - WMI not working
	{
		$CPUName = Get-Reg -Hive LocalMachine -Key $REG_PROCESSOR_NAME -Value "ProcessorNameString"
		$arrCPUNames = $CPUName.Trim()
		if($CPUName -eq "Registry Access Error") { $scanReg = $false }
	}
	else
	{ if($verbose) { Log-MSG -Msg "`tNo CPU Info was collected"  -LogFile $LOG_FILE_PATH }	}

	return $arrCPUNames
}

Function Get-OSDetails
{
<# 
.SYNOPSIS 
	Method to Collect Operation System infomration

.DESCRIPTION
	Method to collect Operation System information using WMI or registry.
	Some of the infomration is available only vai WMI and not Registry
	The Information collected:
		OS Name
		Service Pack
		Last Boot time
		
.PARAMETER (Switch) scanWMI
	Whether to scan using WMI (True / False)
	(Preffered method)
.PARAMETER (Switch) scanReg
	Whether to scan using Registry (True / False)
#>
	param([switch]$scanWMI, [switch]$scanReg)
	# Create an Object With Empty Values
	$OS = "" | Select Caption,CSDVersion,LastBootUpTime
	if($scanWMI) # Scan WMI
	{
		# Collect Operating System and Service Pack Information Using WMI
		$OS = Get-WMIItem -Class "win32_operatingsystem" -Item Caption,CSDVersion,LastBootUpTime
		$OS.Caption = $OS.Caption.Replace(","," ")
		If($OS.LastBootUpTime -ne $Null)
		{ $OS.LastBootUpTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($OS.LastBootUpTime) }
	}
	elseif($scanReg) # Scan Registry
	{
		$OS.Caption = Get-Reg -Hive LocalMachine -Key "SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Value "ProductName"
		if (($OS.Caption -eq "Registry Access Error") -or ($OS.Caption -eq $null)) { $scanReg = $false }
		$OS.CSDVersion = Get-Reg -Hive LocalMachine -Key "SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Value "CSDVersion"
		$OS.LastBootUpTime = "-"
	}
	
	return $OS
}

Function Get-DiskDriveDetails
{
<# 
.SYNOPSIS 
	Method to Collect Disk infomration

.DESCRIPTION
	Method to collect disk information using WMI or registry.
	Registry collection is not supported
	The Information collected:
		Disk Name (Title)
		Disk size (Total / Aviable)
		
.PARAMETER (Switch) scanWMI
	Whether to scan using WMI (True / False)
	(Preffered method)
.PARAMETER (Switch) scanReg
	Whether to scan using Registry (True / False)
#>
	param([switch]$scanWMI, [switch]$scanReg)
	if($scanWMI) # Scan WMI
	{
		# Collect Disk Drive Information Using WMI
		$DriveInfo = Get-WMIItem -Class "Win32_LogicalDisk" -Item Caption,Size,FreeSpace -Filter "DriveType<=3"
		# Format Every Drive Size and Free Space
		foreach($DRSize in $DriveInfo)
		{ # Check Object Size and Format accordingly
			if($DRSize.Size -ne $Null)
			{
				$DRSize.Size = Format-Number $DRSize.Size
				$DRSize.FreeSpace = Format-Number $DRSize.FreeSpace
			}
			else
			{ 
				$DRSize.Size = "-"
				$DRSize.FreeSpace = "-"
			}
		}
		# Disk Drives Can Contain Multiple Values
		$arrDiskDrives = @() 
		foreach($Drive in $DriveInfo){
			$arrDiskDrives += "{0} ({2}/{1})" -f $Drive.Caption, $Drive.Size, $Drive.FreeSpace
			}
	}
	else
	{ if($verbose) { Log-MSG -MSG "`tNo Disk Drive Info collected"  -LogFile $LOG_FILE_PATH } }
	
	return $arrDiskDrives
}

Function Get-TimeZoneDetails
{
<# 
.SYNOPSIS 
	Method to Collect Timezone infomration

.DESCRIPTION
	Method to collect Timezone information using WMI or registry.
	Registry collection is not supported
	The Information collected:
		Standard Name
		Bias (Hours)
		
.PARAMETER (Switch) scanWMI
	Whether to scan using WMI (True / False)
	(Preffered method)
.PARAMETER (Switch) scanReg
	Whether to scan using Registry (True / False)
#>
	param([switch]$scanWMI, [switch]$scanReg)
	if($scanWMI) # Scan WMI
	{
		# Collect Time Zone Information Using WMI
		$TimeZone = Get-WMIItem -Class "Win32_TimeZone" -Item Bias,StandardName
		$TimeZone.Bias = $TimeZone.Bias/60
	}
	else
	{ if($verbose) { Log-MSG -MSG "`tNo Time Zone Data is collected"  -LogFile $LOG_FILE_PATH } }		
	
	return $TimeZone
}

Function Get-OSInfo
{
<#
  .SYNOPSIS
  Collects general Operating system information to a file
  .DESCRIPTION
  Collects Operating system information to a file.
  All the information will be saved in a txt file called "OSInfo.txt" in the destination folder.
  Information collected includes:
	OS Name
	OS Service Pack number
	CPU name
	CPU number of cores
	Total RAM
	Total / Free Disk space
   
  .EXAMPLE
  Get-OSInfo -OutFolder "C:\Path to export folder"
  
  .PARAMETER OutFolder
  The folder where to save the OSInfo.txt file
  
#>
	param
	(
		[Parameter(Mandatory=$true,HelpMessage="Enter the output folder path")]
		[ValidateScript({Test-Path $_})]
		[String]$OutFolder
	)
	
	Log-Msg -MSG "Collecting OS Information for machine $($ENV:ComputerName)" -SubHeader  -LogFile $LOG_FILE_PATH
	$scanReg = $scanWMI = $true # Assume to be True until Check
	# Check Remote Registry Service, Activate it if Stopped
	if( Check-RegistryService -match "Running" ) { $scanReg = $true } else { $scanReg = $false }
	$checkWMI = Get-WMIItem Win32_computersystem -Item Caption
	if(($checkWMI -eq "WMI Error") -or ($checkWMI -eq $Null)) 
	{ 
		$scanWMI = $false  # Disable WMI Scanning
		if($verbose){ LOG-MSG "$strComputer Has WMI Problems - Switching to Registry Scan" -Type "Warning" }
	}
	Log-MSG -MSG "Collecting Basic Computer details and OS Information"  -LogFile $LOG_FILE_PATH
	$_CompDetails = (Get-ComputerDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	$_OSDetails = (Get-OSDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	Log-MSG -MSG "Collecting CPU Information"  -LogFile $LOG_FILE_PATH
	$_CPUDetails = (Get-CPUDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	Log-MSG -MSG "Collecting Disk Information"  -LogFile $LOG_FILE_PATH
	$_DiskDetails = (Get-DiskDriveDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	Log-MSG -MSG "Collecting Timezone Information"  -LogFile $LOG_FILE_PATH
	$_TimezoneDetails = (Get-TimeZoneDetails -scanWMI:$scanWMI -scanReg:$scanReg)

	# Export information to the OSInfo file
	$exportFile = "$OutFolder\OSInfo.txt"
	$machineDetails = @"
Machine Name: {0}
Domain: {1}
System Type: {2} ({3} {4})
Memory: {5}/{6}
User Name: {7}
"@ 
	$machineDetails -f $_CompDetails.Caption,$_CompDetails.Domain,$_CompDetails.SystemType,$_CompDetails.Manufacturer,$_CompDetails.Model,$_CompDetails.AvialableMem,$_CompDetails.TotalPhysicalMemory,$_CompDetails.UserName | Out-File $exportFile
	$osDetails = @"
Operating System: {0} (Service Pack {1})
Last Boot time: {2}
"@
	$osDetails -f $_OSDetails.Caption,$_OSDetails.CSDVersion,$_OSDetails.LastBootUpTime | Out-File $exportFile -Append
	"Time Zone information: {0} ({1})" -f $_TimezoneDetails.StandardName,$_TimezoneDetails.Bias | Out-File $exportFile -Append
	"Number of Cores: $($_CPUDetails.Count)" | Out-File $exportFile -Append
	"CPU Name(s):" | Out-File $exportFile -Append
	ForEach($cpu in $_CPUDetails)
	{
		"`t$cpu" | Out-File $exportFile -Append
	}
	"Disk Drives:" | Out-File $exportFile -Append
	ForEach ($drive in $_DiskDetails)
	{
		"`t$drive" | Out-File $exportFile -Append
	}
	Log-MSG -MSG "Done exporting Basic Computer details and OS Information"  -LogFile $LOG_FILE_PATH
}
#endregion

#region Collect Event Viewer
Function Get-EventViewer
{
<#
  .SYNOPSIS
  Collect Event Log information to a CSV file
  .DESCRIPTION
  Export Event Log logs to a CSV to a CSV file called <LogName>-EventLog.csv in the Output folder specified
  .EXAMPLE
  Get-EventViewer -Log Application -From 1/1/2016 -To 1/1/2017 -OutFolder "C:\temp"
  .PARAMETER Log
  The Event Log Name to export (Application, System)
  .PARAMETER From
  Retrive event log entries from a specific date
  .PARAMETER To
  Retrive event log entries to a specific date
  .PARAMETER OutFolder
  The folder where to save the <LogName>-EventLog.csv file
#>
	param(
		[Parameter(Mandatory=$true,HelpMessage="Enter the output folder path")]
		[ValidateScript({Test-Path $_})]
		[String]$OutFolder,
		[Parameter(Mandatory=$true,HelpMessage="Enter the Name of the Event viewer log")]
		[ValidateSet("Application","System","Security")]
		[String]$Log,
		[Parameter(Mandatory=$false)]
		$From,
		[Parameter(Mandatory=$false)]
		$To		
	)
	If ([string]::IsNullOrEmpty($From) -or [string]::IsNullOrEmpty($To))
	{
		$From = $(Get-Date (Get-Date).AddHours(-72) -Format $g_DateTimePattern)
		$To = $(Get-Date -Format $g_DateTimePattern)
	}
	else
	{
		$From = $(Get-DateFromParameter $From)
		$To = $(Get-DateFromParameter $To)
	}
	$exportFile = "$OutFolder\$Log-EventLog.csv"

	Get-EventLog -LogName $Log -After $From -Before $To -ErrorAction SilentlyContinue | Select EntryType,TimeWritten,Source,EventID,Message,UserName,MachineName, @{name='ReplacementStrings';Expression={ $_.ReplacementStrings -join ';'}} | where {$_.ReplacementStrings -notmatch '^S-1-5'} | Export-Csv $exportFile
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
	
	Log-Msg -MSG "Start running component script" -SubHeader  -LogFile $LOG_FILE_PATH
	If ([string]::IsNullOrEmpty($LogsFrom)) { $LogsFrom = $(Get-Date (Get-Date).AddHours(-72) -Format $g_DateTimePattern) } 
	If ([string]::IsNullOrEmpty($LogsTo)) { $LogsTo = $(Get-Date -Format $g_DateTimePattern) }
	Log-MSG -MSG "Will collect logs from $LogsFrom to $LogsTo" -Type Debug  -LogFile $LOG_FILE_PATH
	$scriptPathAndArgs = "& `"{0}`" -ComponentPath '$Path' -DestFolderPath '$OutFolder' -TimeframeFrom `"$LogsFrom`" -TimeframeTo `"$LogsTo`""
	$scriptExpression = ""
		
	Switch($Name)
	{
		"Vault" { $scriptExpression = ($scriptPathAndArgs -f $VAULT_LOG_COLLECTION_PATH) }
		"CPM" { $scriptExpression = ($scriptPathAndArgs -f $CPM_LOG_COLLECTION_PATH) }
		"PVWA" { $scriptExpression = ($scriptPathAndArgs -f $PVWA_LOG_COLLECTION_PATH) }
		"PSM" { $scriptExpression = ($scriptPathAndArgs -f $PSM_LOG_COLLECTION_PATH)
             
			Log-MSG -MSG "Exporting PSM Security Event Viewer logs from $TimeframeFrom to $TimeframeTo" -Type Debug -LogFile $LOG_FILE_PATH
			Get-EventViewer -Log "Security" -OutFolder $tempFolder -From $TimeframeFrom -To $TimeframeTo
            Log-MSG -MSG "Exporting PSM AppLocker Event Viewer logs from $TimeframeFrom to $TimeframeTo" -Type Debug -LogFile $LOG_FILE_PATH
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
	Log-MSG -MSG "Running $Name Collection script" -Type Debug -LogFile $LOG_FILE_PATH
	Log-MSG -MSG $scriptExpression -Type Debug  -LogFile $LOG_FILE_PATH
	Invoke-Expression $scriptExpression
	Log-Msg -MSG "Done running script" -SubHeader -LogFile $LOG_FILE_PATH
}
#endregion

#---------------
Log-MSG -MSG "Starting script (v$ScriptVersion)" -Header  -LogFile $LOG_FILE_PATH
if($InDebug) { Log-MSG -MSG "Running in Debug Mode"  -LogFile $LOG_FILE_PATH }
if($InVerbose) { Log-MSG -MSG "Running in Verbose Mode"  -LogFile $LOG_FILE_PATH }
Log-MSG -MSG "Running PowerShell version $($PSVersionTable.PSVersion.Major) compatible of versions $($PSVersionTable.PSCompatibleVersions -join ", ")"  -LogFile $LOG_FILE_PATH

If($Analyze)
{
	Throw "This is not yet supported"
	exit
}

If($Collect)
{
	# Write the CYBRSupport Tool Version file
	"CYBRSupportTool Running version $ScriptVersion" | Out-File -FilePath "$ScriptLocation\CYBRSupportTool_Version.txt" -Force
	Log-Msg -MSG "Collecting $ComponentName relevant information"  -LogFile $LOG_FILE_PATH
	
	$zipFileName = "$ComponentName-$(Get-Date -Format 'ddMMyyyy_hhmmss')"
	$tempFolder = New-Item -Type "Directory" -Path "$($ENV:TEMP)\$zipFileName"
		
	If([string]::IsNullOrEmpty($TimeframeFrom)) { $TimeframeFrom = $(Get-Date (Get-Date).AddHours(-72) -Format $g_DateTimePattern) }
	If([string]::IsNullOrEmpty($TimeframeTo)) { $TimeframeTo = $(Get-Date -Format $g_DateTimePattern) }

	# Get OS Info
	If($OSInfo) { Get-OSInfo -OutFolder $tempFolder }
	# Get Event Viewer
	If($EventViewer) 
	{ 
		Log-MSG -MSG "Exporting Application Event Viewer logs from $TimeframeFrom to $TimeframeTo"  -LogFile $LOG_FILE_PATH
		Get-EventViewer -Log "Application" -OutFolder $tempFolder -From $TimeframeFrom -To $TimeframeTo
        Log-MSG -MSG "Exporting SystemEvent Viewer logs from $TimeframeFrom to $TimeframeTo"  -LogFile $LOG_FILE_PATH
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
		Log-Msg -MSG "Zipping files from $tempFolder to $ZipLocationFolder\$zipFileName.zip"  -LogFile $LOG_FILE_PATH
		Zip-Files -zipFileName "$ZipLocationFolder\$zipFileName.zip" -sourceDir $tempFolder
		Log-Msg -MSG "Zip file size $(Format-Number (Get-Item "$ZipLocationFolder\$zipFileName.zip").Length)"  -LogFile $LOG_FILE_PATH
		# Delete the Temp folder
		Log-Msg -MSG "Deleting temp folder $tempFolder"  -LogFile $LOG_FILE_PATH
		Remove-Item -Recurse -Force -Path $tempFolder
	}
	Else
	{
		Throw "$ComponentName installation path ($InstallPath) does not exit"
		Exit
	}
}

Log-MSG -MSG "Script ended" -Footer  -LogFile $LOG_FILE_PATH