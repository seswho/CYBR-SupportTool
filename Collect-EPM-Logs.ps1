###########################################################################
#
# SCRIPT NAME: Collect EPM Logs
#
# VERSION HISTORY:
# 1.0 12/09/2017 - Initial release
# 1.1 05/11/2017 - Fixed EPM Log locations and added MSSQL version
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk EPM Server logs

.DESCRIPTION
            This script can collect logs and configurations of an EPM Server
			
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

Write-LogMessage -MSG "Collecting EPM Files" -Type Debug
Write-LogMessage -MSG "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrEPMFilePaths = @()

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting EPM file versions and additional information" -Type Debug
$EPMVersions = "$DestFolderPath\_EPMFileVersions.txt"
"EPM: "+$(Get-FileVersion "$ComponentPath\VfBackgroundWorker.exe") | Out-File $EPMVersions
$EPMLogFolder = "$ComponentPath\Log"
$EPMTraceFolder = "$($ComponentPath.Replace("VFSVC","PASERVER"))\Trace"
"Log Folder: $EPMLogFolder" | Out-File $EPMVersions -append
"Log Folder: $EPMTraceFolder" | Out-File $EPMVersions -append
$MSSQL = (Get-WMIObject -Query 'Select * from win32_service where Name like "MSSQLSERVER"') | select @{Name="Path"; Expression={$_.PathName.split('"')[1]}} 
If (Test-Path $MSSQL.Path)
{
	"MSSQL: "+$(Get-FileVersion $MSSQL.Path) | Out-File $EPMVersions
}

Write-LogMessage -MSG "Collecting EPM files list by timeframe" -Type Debug
$arrEPMFilePaths += $EPMVersions
# Check that the logs folder is not empty
If(([string]::IsNullOrEmpty($EPMLogFolder) -ne $true) -and (Test-Path $EPMLogFolder))
{
	$arrEPMFilePaths += (Get-FilesByTimeframe -Path "$EPMLogFolder\*.log" -From $TimeframeFrom -To $TimeframeTo)
	$arrEPMFilePaths += (Get-FilesByTimeframe -Path "$EPMLogFolder\*.csv" -From $TimeframeFrom -To $TimeframeTo)
}
else
{
	Write-LogMessage -MSG "EPM Logs folder returned empty" -Error -Type Debug 
}
# Check that the trace folder is not empty
If(([string]::IsNullOrEmpty($EPMTraceFolder) -ne $true) -and (Test-Path $EPMTraceFolder))
{
	$arrEPMFilePaths += (Get-FilesByTimeframe -Path "$EPMTraceFolder\*.txt*" -From $TimeframeFrom -To $TimeframeTo)
	$arrEPMFilePaths += (Get-FilesByTimeframe -Path "$EPMTraceFolder\*.csv" -From $TimeframeFrom -To $TimeframeTo)
}
else
{
	Write-LogMessage -MSG "EPM trace folder returned empty" -Error -Type Debug 
}

Collect-Files -arrFilesPath $arrEPMFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting EPM Files" -Type Debug