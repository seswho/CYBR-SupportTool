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
# 1.11 08/09/2020 - Refactored using shared module
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
$global:LOG_FILE_PATH = "$ScriptLocation\SupportTool_Detect.log"
# Set Module Path
$MODULE_SHARED = "$ScriptLocation\bin\CYBRSupportTool-Shared.psd1"

# Set Debug / Verbose script modes
$InDebug = $PSBoundParameters.Debug.IsPresent
$InVerbose = $PSBoundParameters.Verbose.IsPresent

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
#endregion

#---------------
# Load all relevant modules
$moduleInfos = Load-Modules

# Check if Vault is installed
$Vault = Get-ComponentDetails -ComponentName "Vault" -ComponentServicePath $REGKEY_VAULTSERVICE_OLD -ComponentExecutable "dbmain.exe"
If($Vault -ne $null) { Write-Host ("{0}|{1}" -f $Vault.Name, $Vault.Path) }
# Check if Vault DR is installed
$VaultDR = Get-ComponentDetails -ComponentAlias "PADR" -ComponentName "Vault Disaster Recovery" -ComponentServicePath $REGKEY_DRSERVICE -ComponentExecutable "PADR.exe"
If($VaultDR -ne $null) { Write-Host ("{0}|{1}" -f $VaultDR.Name, $VaultDR.Path) }
# Check if Cluster Vault Manager is installed
$CVM = Get-ComponentDetails -ComponentAlias "CVM" -ComponentName "Cluster Vault Manager" -ComponentServicePath $REGKEY_CVMSERVICE -ComponentExecutable "ClusterVault.exe"
If($CVM -ne $null) { Write-Host ("{0}|{1}" -f $CVM.Name, $CVM.Path) }
# Check if RabbitMQ is installed Distributed Vaults for PSM/PVWA
$MQ = Get-ComponentDetails -ComponentAlias "RabbitMQ" -ComponentName "RabbitMQ for Distributed Vaults" -ComponentServicePath $REGKEY_RMQSERVICE -ComponentExecutable "erlsrv.exe"
If($MQ -ne $null) { Write-Host ("{0}|{1}" -f $MQ.Name, $MQ.Path) }
# Check if CPM is installed
$CPM = Get-ComponentDetails -ComponentName "CPM" -ComponentServicePath $REGKEY_CPMSERVICE_OLD -ComponentExecutable "PMEngine.exe"
If($CPM -ne $null) { Write-Host ("{0}|{1}" -f $CPM.Name, $CPM.Path) }
# Check if PVWA is installed
$PVWA = Get-ComponentDetails -ComponentName "PVWA" -ComponentServicePath ($REGKEY_PVWASERVICE) -ComponentExecutable "Services\\CyberArkScheduledTasks.exe"
If($PVWA -ne $null) { Write-Host ("{0}|{1}" -f $PVWA.Name, $PVWA.Path) }
# Check if PSM is installed
$PSM = Get-ComponentDetails -ComponentName "PSM" -ComponentServicePath ($REGKEY_PSMSERVICE) -ComponentExecutable "CAPSM.exe"
If($PSM -ne $null) { Write-Host ("{0}|{1}" -f $PSM.Name, $PSM.Path) }
# Check if AIM is installed
$AIM = Get-ComponentDetails -ComponentName "AIM" -ComponentServicePath ($REGKEY_AIMSERVICE) -ComponentExecutable "AppProvider.exe"
If($AIM -ne $null) { Write-Host ("{0}|{1}" -f $AIM.Name, $AIM.Path) }
# Check if EPM Server is installed
$EPM = Get-ComponentDetails -ComponentName "EPM" -ComponentServicePath ($REGKEY_EPMSERVICE) -ComponentExecutable "VfBackgroundWorker.exe"
If($EPM -ne $null) { Write-Host ("{0}|{1}" -f $EPM.Name, $EPM.Path) }

# UnLoad loaded modules
UnLoad-Modules $moduleInfos