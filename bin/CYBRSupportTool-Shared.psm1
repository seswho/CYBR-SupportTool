#region Writer Functions
# @FUNCTION@ ======================================================================================================================
# Name...........: Write-LogMessage
# Description....: Writes the message to log and screen
# Parameters.....: LogFile, MSG, (Switch)Header, (Switch)SubHeader, (Switch)Footer, Type
# Return Values..: None
# =================================================================================================================================
Function Write-LogMessage
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
	The type of the message to log (Info, Warning, Error, Debug, Verbose , Success, LogOnly)
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
		[ValidateSet("Info","Warning","Error","Debug","Verbose", "Success", "LogOnly")]
		[String]$type = "Info",
		[Parameter(Mandatory=$false)]
		[String]$LogFile = $LOG_FILE_PATH
	)
	Try{
		If ($Header) {
			"=======================================" | Out-File -Append -FilePath $LogFile 
			Write-Host "=======================================" -ForegroundColor Magenta
		}
		ElseIf($SubHeader) { 
			"------------------------------------" | Out-File -Append -FilePath $LogFile 
			Write-Host "------------------------------------" -ForegroundColor Magenta
		}
		
		$msgToWrite = "[$(Get-Date -Format "yyyy-MM-dd hh:mm:ss")]`t"
		$writeToFile = $true
		# Replace empty message with 'N/A'
		if([string]::IsNullOrEmpty($Msg)) { $Msg = "N/A" }
		
		# Mask Passwords
		if($Msg -match '((?:password|credentials|secret)\s{0,}["\:=]{1,}\s{0,}["]{0,})(?=([\w`~!@#$%^&*()-_\=\+\\\/|;:\.,\[\]{}]+))')
		{
			$Msg = $Msg.Replace($Matches[2],"****")
		}
		# Check the message type
		switch ($type)
		{
			{($_ -eq "Info") -or ($_ -eq "LogOnly")} 
			{ 
				If($_ -eq "Info")
				{
					Write-Host $MSG.ToString() -ForegroundColor $(If($Header -or $SubHeader) { "Magenta" } Else { "White" })
				}
				$msgToWrite += "[INFO]`t$Msg"
			}
			"Success" { 
				Write-Host $MSG.ToString() -ForegroundColor Green
				$msgToWrite += "[SUCCESS]`t$Msg"
			}
			"Warning" {
				Write-Host $MSG.ToString() -ForegroundColor Yellow
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
					Write-Verbose -Msg $MSG
					$msgToWrite += "[VERBOSE]`t$Msg"
				}
				else { $writeToFile = $False }
			}
		}

		If($writeToFile) { $msgToWrite | Out-File -Append -FilePath $LogFile }
		If ($Footer) { 
			"=======================================" | Out-File -Append -FilePath $LogFile 
			Write-Host "=======================================" -ForegroundColor Magenta
		}
	}
	catch{
		Throw $(New-Object System.Exception ("Cannot write message"),$_.Exception)
	}
}
Export-ModuleMember -Function Write-LogMessage

# @FUNCTION@ ======================================================================================================================
# Name...........: Join-ExceptionMessage
# Description....: Formats exception messages
# Parameters.....: Exception
# Return Values..: Formatted String of Exception messages
# =================================================================================================================================
Function Join-ExceptionMessage
{
<# 
.SYNOPSIS 
	Formats exception messages
.DESCRIPTION
	Formats exception messages
.PARAMETER Exception
	The Exception object to format
#>
	param(
		[Exception]$e
	)

	Begin {
	}
	Process {
		$msg = "Source:{0}; Message: {1}" -f $e.Source, $e.Message
		while ($e.InnerException) {
		  $e = $e.InnerException
		  $msg += "`n`t->Source:{0}; Message: {1}" -f $e.Source, $e.Message
		}
		return $msg
	}
	End {
	}
}
Export-ModuleMember -Function Join-ExceptionMessage
#endregion

#region Helper Functions
#region Internal Methods
# @FUNCTION@ ======================================================================================================================
# Name...........: Get-ServiceInstallPath
# Description....: Get the installation path of a service
# Parameters.....: Service Name
# Return Values..: $true
#                  $false
# =================================================================================================================================
# Save the Services List
$m_ServiceList = $null
Function Get-ServiceInstallPath
{
<#
  .SYNOPSIS
  Get the installation path of a service
  .DESCRIPTION
  The function receive the service name and return the path or returns NULL if not found
  .EXAMPLE
  (Get-ServiceInstallPath $<ServiceName>) -ne $NULL
  .PARAMETER ServiceName
  The service name to query. Just one.
 #>
	param (
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String]$ServiceName
	)
	Begin {
		
	}
	Process {
		$retInstallPath = $Null
		try{
			if ($m_ServiceList -eq $null)
			{
				$m_ServiceList = Get-ChildItem "HKLM:\System\CurrentControlSet\Services" | ForEach-Object { Get-ItemProperty $_.pspath }
				#$m_ServiceList = Get-Reg -Hive "LocalMachine" -Key System\CurrentControlSet\Services -Value $null
			}
			$regPath =  $m_ServiceList | Where-Object {$_.PSChildName -eq $ServiceName}
			If ($regPath -ne $Null)
			{
				$retInstallPath = $regPath.ImagePath.Substring($regPath.ImagePath.IndexOf('"'),$regPath.ImagePath.LastIndexOf('"')+1)
			}	
		} 
		catch{
			Throw $(New-Object System.Exception ("Cannot get Service Install path for $ServiceName",$_.Exception))
		}
		
		return $retInstallPath
	}
	End {
		
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
	Write-LogMessage -Type "Debug" -Msg "$ServiceName Service Status is: $($svcStatus.Status)"  -LogFile $LOG_FILE_PATH
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
	trap [Exception] { if($verbose){ Write-LogMessage -Msg "$RemoteComputer Has Registry Problems:`n`t$_" -Type "Error"   -LogFile $LOG_FILE_PATH }; continue; }	
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
#endregion

#region Files Methods
# @FUNCTION@ ======================================================================================================================
# Name...........: Get-FileVersion
# Description....: Method to return a file version
# Parameters.....: File Path
# Return Values..: File version
# =================================================================================================================================
Function Get-FileVersion
{
<# 
.SYNOPSIS 
	Method to return a file version

.DESCRIPTION
	Returns the File version and Build number
	Returns Null if not found

.PARAMETER FilePath
	The path to the file to query
#>
	param (
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String]$filePath
	)
	$retFileVersion = $Null
	try{
		If(Test-Path $filePath)
		{
			$path = Resolve-Path $filePath 
			$retFileVersion = ($path | Get-Item | select VersionInfo).VersionInfo.ProductVersion
		}
					
		return $retFileVersion
	}
	catch{
		Throw $(New-Object System.Exception ("Get-FileVersion: Cannot get File ($filePath) version",$_.Exception))
	}
}
Export-ModuleMember -Function Get-FileVersion

# @FUNCTION@ ======================================================================================================================
# Name...........: Get-FilePath
# Description....: Method to return the full file path
# Parameters.....: File Path
# Return Values..: File Path
# =================================================================================================================================
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
	param(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String]$filePath
	)
	
	try
	{
		If(Test-Path $filePath)
		{
			Write-LogMessage -Type Debug -Msg "Retrieving file: $filePath"  -LogFile $LOG_FILE_PATH
			return (Resolve-Path $filePath)
		}
		else
		{
			Write-LogMessage -Type Error -Msg "File '$filePath' Not found"  -LogFile $LOG_FILE_PATH
			return $null
		}
	}
	catch
	{
		Throw $(New-Object System.Exception ("Get-FilePath: Cannot resolve File ($filePath) path",$_.Exception))
	}
}
Export-ModuleMember -Function Get-FilePath

# @FUNCTION@ ======================================================================================================================
# Name...........: Get-FilesByTimeframe
# Description....: Method to collect all files that were modified in a specific time frame
# Parameters.....: File Path
# Return Values..: File Path
# =================================================================================================================================
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
	param (
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String]$path, 
		$from, 
		$to
	)
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
				Write-LogMessage "Adding File $($file.FullName) to Files collection" -Type Debug  -LogFile $LOG_FILE_PATH
				$arrFilePaths += (Get-FilePath $file.FullName)
			}
			return $arrFilePaths
		}
		else
		{
			Write-LogMessage -Type Error -Msg "The Path received ($path) to collect files by time frame (From: $from; To: $to;) was empty or does not exist"  -LogFile $LOG_FILE_PATH
		}
	}
	catch
	{
		Throw $(New-Object System.Exception ("Get-FilesByTimeframe: Cannot get Files by time frame",$_.Exception))
	}
}
Export-ModuleMember -Function Get-FilesByTimeframe

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
		
	Write-LogMessage -Msg "Collecting Files to $destFolder..."  -LogFile $LOG_FILE_PATH
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
				Write-LogMessage -Type Error -Msg "The Path received ($filePath) to collect files was empty or does not exist"  -LogFile $LOG_FILE_PATH
			}
		}
		
		Write-LogMessage "Completed collecting files!"  -LogFile $LOG_FILE_PATH
	}
	catch
	{
		Throw $(New-Object System.Exception ("Collect-Files: Cannot collect files",$_.Exception))
	}
}
Export-ModuleMember -Function Collect-Files
#endregion

#region OS Environment information Methods
# Set Registry Keys
$REG_COMPUTER_NAME = "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" # \ComputerName
$REG_PROCESSOR_NAME = "HARDWARE\DESCRIPTION\System\CentralProcessor\0" # \ProcessorNameString
$REG_HKLM_ENIVORNMENT = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment" # \NUM_OF_PROCESSOR, \PROCESSOR_ARCHITECTURE
$REG_USER_NAME = "Software\Microsoft\Windows\CurrentVersion\Explorer" # HKCU \Logon User Name
$REG_USER_DOMAIN = "SOFTWARE\Policies\Microsoft\System\DNSClient" # PrimaryDnsSuffix
$REG_HKCU_ENIVORNMENT = "Volatile Environment" # \USERDNSDOMAIN
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
		if($verbose) { Write-LogMessage -Type Warning -Msg "`tSome of the Computers Details will not be collected" }
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
	{ if($verbose) { Write-LogMessage -Msg "`tNo CPU Info was collected"  -LogFile $LOG_FILE_PATH }	}

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
	{ if($verbose) { Write-LogMessage -MSG "`tNo Disk Drive Info collected"  -LogFile $LOG_FILE_PATH } }
	
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
	{ if($verbose) { Write-LogMessage -MSG "`tNo Time Zone Data is collected"  -LogFile $LOG_FILE_PATH } }		
	
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
	
	Write-LogMessage -MSG "Collecting OS Information for machine $($ENV:ComputerName)" -SubHeader  -LogFile $LOG_FILE_PATH
	$scanReg = $scanWMI = $true # Assume to be True until Check
	# Check Remote Registry Service, Activate it if Stopped
	if( Check-RegistryService -match "Running" ) { $scanReg = $true } else { $scanReg = $false }
	$checkWMI = Get-WMIItem Win32_computersystem -Item Caption
	if(($checkWMI -eq "WMI Error") -or ($checkWMI -eq $Null)) 
	{ 
		$scanWMI = $false  # Disable WMI Scanning
		if($verbose){ Write-LogMessage "$strComputer Has WMI Problems - Switching to Registry Scan" -Type "Warning" }
	}
	Write-LogMessage -MSG "Collecting Basic Computer details and OS Information"  -LogFile $LOG_FILE_PATH
	$_CompDetails = (Get-ComputerDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	$_OSDetails = (Get-OSDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	Write-LogMessage -MSG "Collecting CPU Information"  -LogFile $LOG_FILE_PATH
	$_CPUDetails = (Get-CPUDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	Write-LogMessage -MSG "Collecting Disk Information"  -LogFile $LOG_FILE_PATH
	$_DiskDetails = (Get-DiskDriveDetails -scanWMI:$scanWMI -scanReg:$scanReg)
	Write-LogMessage -MSG "Collecting Timezone Information"  -LogFile $LOG_FILE_PATH
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
	Write-LogMessage -MSG "Done exporting Basic Computer details and OS Information"  -LogFile $LOG_FILE_PATH
}
Export-ModuleMember -Function Get-OSInfo
#endregion

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
Export-ModuleMember -Function Format-Number

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
Export-ModuleMember -Function Get-EventViewer

# @FUNCTION@ ======================================================================================================================
# Name...........: Test-AdminUser
# Description....: Check if the user is a Local Admin
# Parameters.....: None
# Return Values..: True/False
# =================================================================================================================================
Function Test-AdminUser()
{
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    return (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.SecurityIdentifier] "S-1-5-32-544")  # Local Administrators group SID
}
Export-ModuleMember -Function Test-AdminUser

# @FUNCTION@ ======================================================================================================================
# Name...........: Detect-Components
# Description....: Detects all CyberArk Components installed on the local server
# Parameters.....: None
# Return Values..: Array of detected components on the local server
# =================================================================================================================================
Function Get-ComponentDetails
{
	param (
		[Parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String]$ComponentName, 
		[Parameter(Mandatory=$false)]
		[String]$ComponentAlias = $ComponentName,		
		[Parameter(Mandatory=$true)]
		[String[]]$ComponentServicePath, 
		[Parameter(Mandatory=$true)]
		[String]$ComponentExecutable
	)
	
	try{
		# Check if Vault DR is installed
		Write-LogMessage -Type "Debug" -MSG "Searching for $ComponentName..."
		$componentPath = $null
		ForEach($path in $ComponentServicePath)
		{
			if(($componentPath = $(Get-ServiceInstallPath $path)) -ne $NULL)
			{ break }
		}
		if($componentPath -ne $null)
		{
			Write-LogMessage -Type Success -MSG "Found $ComponentName"
			$escapedEXE = $ComponentExecutable.Replace('.',"\.")
			If($componentPath -match "([\w\s:()\\]{1,})(?=$escapedEXE)")
			{
				$componentPath = $Matches[0]
			}
			$fileVersion = Get-FileVersion "$componentPath\$ComponentExecutable"
			return New-Object PSObject -Property @{Name=$ComponentAlias;Path=$componentPath;Version=$fileVersion}
		}
		else
		{
			return $null
		}
	} catch {
		Write-LogMessage -Type "Error" -Msg "Error detecting '$ComponentName' component. Error: $(Join-ExceptionMessage $_.Exception)"
	}
}
Export-ModuleMember -Function Get-ComponentDetails