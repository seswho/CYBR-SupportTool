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
