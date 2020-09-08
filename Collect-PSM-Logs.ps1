###########################################################################
#
# SCRIPT NAME: Collect PSM Logs
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.1 06/04/2107 - Update Logs and comments
# 1.2 09/08/2107 - Change From/To Type to string
# 1.3 14/08/2107 - Collecting more files
# 1.4 29/08/2107 - Minor bug fixes
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk PSM server logs

.DESCRIPTION
            This script can collect logs and configurations of a PSM server
			
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

Write-LogMessage -MSG "Collecting PSM Files" -Type Debug
Write-LogMessage -MSG "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrPSMFilePaths = @()

# Parse the PSM INI configuration file
Write-LogMessage -MSG "Parsing PSM configuration file ($ComponentPath\basic_psm.ini)" -Type Debug
$psmINI = Get-IniContent "$ComponentPath\basic_psm.ini"
$PSMLogFolder = (Get-FilePath $psmINI["Main"]["LogsFolder"].Replace('"',""))
$PSMVaultPath = (Get-FilePath $psmINI["Main"]["PSMVaultFile"].Replace('"',""))

# Create a file with the relevant file versions
Write-LogMessage -MSG "Collecting PSM file versions and additional information" -Type Debug
$PSMVersions = "$DestFolderPath\_PSMFileVersions.txt"
"PSM: "+$(Get-FileVersion "$ComponentPath\CAPSM.exe") | Out-File $PSMVersions
"Configuration Safe: "+$($psmINI["Main"]["ConfigurationSafe"]) | Out-File $PSMVersions -append
"PSM server ID: "+$($psmINI["Main"]["PSMServerId"]) | Out-File $PSMVersions -append
"PSM server Admin ID: "+$($psmINI["Main"]["PSMServerAdminId"]) | Out-File $PSMVersions -append
"Log Folder: $PSMLogFolder" | Out-File $PSMVersions -append

$arrPSMFilePaths += $PSMVersions

#
# Collect the files from the Hardening folder
Write-LogMessage -MSG "Collecting PSM hardening files" -Type Debug
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\Hardening\*.log")
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\Hardening\*.csv")
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\Hardening\PSMConfigureAppLocker.*")
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\Hardening\PSMHardening.ps1")

# Check that the logs folder is not empty
Write-LogMessage -MSG "Collecting PSM files list by timeframe" -Type Debug
If(([string]::IsNullOrEmpty($PSMLogFolder) -ne $true) -and (Test-Path $PSMLogFolder))
{
	$arrPSMFilePaths += (Get-FilePath "$PSMLogFolder\PSMConsole.log")
	$arrPSMFilePaths += (Get-FilePath "$PSMLogFolder\PSMTrace.log")
	$arrPSMFilePaths += (Get-FilesByTimeframe -Path "$PSMLogFolder\Old\*.log" -From $TimeframeFrom -To $TimeframeTo)
	$arrPSMFilePaths += (Get-FilesByTimeframe -Path "$PSMLogFolder\Components\*.log" -From $TimeframeFrom -To $TimeframeTo)
	$arrPSMFilePaths += (Get-FilesByTimeframe -Path "$PSMLogFolder\Components\Old\*.log" -From $TimeframeFrom -To $TimeframeTo)
}
else
{
	Write-LogMessage -MSG "PSM Logs folder returned empty" -Error -Type Debug 
}
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\Temp\Policies.xml")
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\Temp\PVConfiguration.xml")
$arrPSMFilePaths += (Get-FilePath "$ComponentPath\basic_psm.ini")
$arrPSMFilePaths += (Get-FilePath $PSMVaultPath)

Collect-Files -arrFilesPath $arrPSMFilePaths -destFolder $DestFolderPath
Write-LogMessage -MSG "Done Collecting PSM Files" -Type Debug