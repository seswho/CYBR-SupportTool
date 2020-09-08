###########################################################################
#
# SCRIPT NAME: Collect AIM Logs
#
# VERSION HISTORY:
# 1.0 03/09/2017 - Initial release
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk AIM (PIM Provider) logs

.DESCRIPTION
            This script can collect logs and configurations of an AIM (PIM Provider) provider
			
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
	[Parameter(Mandatory=$true,HelpMessage="Enter the Component installation path")]
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

Function Get-IniContent 
{
<# 
.SYNOPSIS 
	Method to query settings from an INI file

.DESCRIPTION
	This method will collect settings from an INI configuration file.
	
.PARAMETER filePath
	The INI file path to analyse
#>
	param ($filePath)
    $ini = @{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = “Comment” + $CommentCount
            $ini[$section][$name] = $value
        } 
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

Write-LogMessage -MSG "Collecting AIM Files" -Type Debug
Write-LogMessage -MSG "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrAIMFilePaths = @()

# Parse the AIM INI configuration file
Write-LogMessage -MSG "Parsing AIM configuration file ($ComponentPath\basic_appprovider.conf)" -Type Debug
$AIMConf = Get-IniContent "$ComponentPath\basic_appprovider.conf"
$AIMLogFolder = (Get-FilePath $AIMConf["Main"]["LogsFolder"].Replace('"',""))
$AIMVaultPath = (Get-FilePath $AIMConf["Main"]["AppProviderVaultFile"].Replace('"',""))

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting AIM file versions and additional information" -Type Debug
$AIMVersions = "$DestFolderPath\_AIMFileVersions.txt"
"AIM: "+$(Get-FileVersion "$ComponentPath\AppProvider.exe") | Out-File $AIMVersions
"Configuration Safe: "+$($AIMConf["Main"]["PIMConfigurationSafe"]) | Out-File $AIMVersions -append
"Parameters Safe: "+$($AIMConf["Main"]["AppProviderParmsSafe"]) | Out-File $AIMVersions -append
"Parameters File: "+$($AIMConf["Main"]["AppProviderVaultParmsFile"]) | Out-File $AIMVersions -append
"Log Folder: $AIMLogFolder" | Out-File $AIMVersions -append

Write-LogMessage -MSG "Collecting AIM files list by timeframe" -Type Debug
$arrAIMFilePaths += $AIMVersions
# Check that the logs folder is not empty
If(([string]::IsNullOrEmpty($AIMLogFolder) -ne $true) -and (Test-Path $AIMLogFolder))
{
	$arrAIMFilePaths += (Get-FilePath "$AIMLogFolder\APPConsole.log")
	$arrAIMFilePaths += (Get-FilePath "$AIMLogFolder\APPTrace.log")
	$arrAIMFilePaths += (Get-FilePath "$AIMLogFolder\APPAudit.log")
	$arrAIMFilePaths += (Get-FilesByTimeframe -Path "$AIMLogFolder\Old\*.log" -From $TimeframeFrom -To $TimeframeTo)
}
else
{
	Write-LogMessage -MSG "AIM Logs folder returned empty" -Error -Type Debug 
}
$arrAIMFilePaths += (Get-FilePath "$ComponentPath\basic_appprovider.conf")
$arrAIMFilePaths += (Get-FilePath $AIMVaultPath)

Collect-Files -arrFilesPath $arrAIMFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting AIM Files" -Type Debug