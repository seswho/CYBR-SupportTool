###########################################################################
#
# SCRIPT NAME: Collect Vault Logs
#
# VERSION HISTORY:
# 1.0 05/04/2017 - Initial release
# 1.1 06/04/2107 - Update Logs and comments
# 1.2 09/08/2107 - Change From/To Type to string
# 1.3 14/08/2107 - Adding more logs to collect
# 1.4 30/01/2019 - Added logic to determine where to find the logs
#				   depending on dbmain.exe version
# 1.5 12/02/2020 - Added the ability to run cavaultmanager diagnosedbreport 
#				   and capture the output to a file 3 times with a 30 seconds
#				   pause between each report. Also added the capture of the
#                  paragent files
# 1.6 26/02/2020 - Add more config and log files to be collected: ENE, passparm.ini, vault.ini, proper log file names for Logic Container (pre-10.5)
#
###########################################################################

<# 
.SYNOPSIS 
            A script to collect the CyberArk Vault server logs

.DESCRIPTION
            This script can collect logs and configurations of a Vault server
			
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

#
# include the shared functions script
. .\CYBRSupportTool-Shared.ps1

#region Helper Functions
Function Get-CertificateChain
{
	param (
		[Parameter(Mandatory=$true,HelpMessage="Enter the target address")]
		[Alias("IP")]
		[String]$Server,
	
		[Parameter(Mandatory=$false,HelpMessage="Enter the target port")]
		[int]$Port = 636
	)
			
	$tcpClient = New-Object Net.Sockets.TcpClient($Server,$Port)
	$tcpClient.ReceiveTimeout = $tcpClient.SendTimeout = 2000;
    
    If($tcpClient.Connected)
    {
		try
		{
			Write-LogMessage -Type Info -LogFile $ldapsCertificateLog -MSG "Getting Certificate Chain for server $Server"
			Write-LogMessage -Type Debug -LogFile $ldapsCertificateLog -MSG "Port $port is operational to server $Server"
			$fnValidateServerCertificate = ${Function:\ValidateServerCertificate}

			$objSslStream = New-Object Net.Security.SslStream($($tcpClient.GetStream()), $False, $fnValidateServerCertificate)
			$objSslStream.AuthenticateAsClient($Server)
			
			Write-LogMessage -Type Info -LogFile $ldapsCertificateLog -Msg "Chain verified successfully"
		}
		catch
		{
			Switch ($_.Exception.HResult)
			{
				-2146232800 { Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg "Check that the target LDAPS Host has an appropriate, valid, Certificate and the appropriate port for LDAPS was defined." }
				-2147467259 { Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg "Check network communication between local host and target LDAPS Host. (IE Firewalls, Routing, Proxies, Port Assignment, Name Resolution, etc.)" }
				-2146233087 { Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg "Network Communications successfully established, but could not validate Certificate or Certificate Chain due to the above Certificate Verification Errors." }
			}
			Return
        }
    }
    Else
    {
        Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -MSG "Port $port on $Server is closed, You may need to contact your IT team to open it."  
    }
}

Function Check-CertificateChain 
{
        Param (
            [Parameter(Mandatory=$true)]
            [string]$Server,

            [Parameter(Mandatory=$true)]
            [string]$Port
        )
        Write-LogMessage -Type Info -LogFile $ldapsCertificateLog -MSG "Opening network communication over Port: $Port in local firewall"
        New-NetFirewallRule -DisplayName "Temp_Allow_LDAPS_Delete_this_rule_if_found" -Profile "Any" -Direction Outbound -Action Allow -Protocol TCP -RemotePort $Port | Out-Null

        Write-LogMessage -Type Info -MSG "Retrieving Certificate Chain"
        Get-CertificateChain -server $Server -port $Port

        Write-LogMessage -Type Info -LogFile $ldapsCertificateLog -MSG "Closing network communication over Port: $Port in local firewall"
        Remove-NetFirewallRule -DisplayName "Temp_Allow_LDAPS_Delete_this_rule_if_found"
}

Function ValidateServerCertificate
{
	param 
	(
		[System.Object]$sender, 
		[Security.Cryptography.X509Certificates.X509Certificate]$certificate, 
		[Security.Cryptography.X509Certificates.X509Chain]$Chain, 
		[Net.Security.SslPolicyErrors]$SslPolicyErrors
	)
	
	# Output data about the certificates in the chain to the Main form console
	$ChainEnum = $Chain.ChainElements.GetEnumerator()
    Write-LogMessage -Type Info -Msg "Chain Elements found: $($Chain.ChainElements.Count)"
	Write-LogMessage -Type Info -Msg "Chain Begin" -SubHeader
	For ($i=1; $i -le $Chain.ChainElements.Count; $i++) {
		Write-LogMessage -Type Debug -Msg "Getting details of Element: $i"
		$ChainEnum.MoveNext()
		$IsRoot = If ($ChainEnum.Current.Certificate.Extensions | Where-Object {$_.OID.FriendlyName -eq 'Authority Key Identifier'}){$false}Else{$true}
		$CertInfoObj = [pscustomobject]@{
			System = "Certificate Chain Cert: $i";
			Root = $IsRoot
			Subject = $ChainEnum.Current.Certificate.Subject;
			SubjectAlternativeName = ($ChainEnum.Current.Certificate.Extensions | Where-Object {$_.OID.FriendlyName -eq 'Subject Alternative Name'}) | ForEach-Object {$_.Format(1)}
			EffectiveDate = $ChainEnum.Current.Certificate.NotBefore;
			ExpirationDate = $ChainEnum.Current.Certificate.NotAfter;
		}
        
		$CertInfoObj | FL | Out-String | Write-LogMessage -Type Info -LogFile $ldapsCertificateLog
	}
	Write-LogMessage -Type Info -LogFile $ldapsCertificateLogo -Msg "Checking Chain StatusInformation..."
	If ($Chain.ChainStatus.StatusInformation) {
		Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg ($Chain.ChainStatus.StatusInformation -join '')
	}
	Else { Write-LogMessage -Type Info -LogFile $ldapsCertificateLog -Msg "No errors found" }
	
	Write-LogMessage -Type Info -LogFile $ldapsCertificateLog -Msg "Chain Ends" -SubHeader

	# Check if the SslPolicy has errors, evaluate, and output to log
	If ($SslPolicyErrors)
	{
		$intSslErrorBitwise = $SslPolicyErrors
		If ($intSslErrorBitwise -ge 4) { 
			$intSslErrorBitwise -= 4
			Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg "Certificate Verification Error: Remote Certificate Chain Errors. Possibly Root Certificate Authority is not trusted on this system."
		}
		If ($intSslErrorBitwise -ge 2) {
			$intSslErrorBitwise -= 2
			Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg "Certificate Verification Error: Remote Certificate Name Mismatch. Check to ensure defined LDAPS Host matches Subject or Subject Alternative Name in remote LDAPS Host's Certificate."
		}
		If ($intSslErrorBitwise -ge 1) {
			$intSslErrorBitwise -= 1
			Write-LogMessage -Type Error -LogFile $ldapsCertificateLog -Msg "Certificate Verification Error: Remote Certificate not Available."
		}
		Return $False
	}
	Return $True
}

Function Get-HostsRecords
{
	$hostsPath = "$env:windir\System32\drivers\etc\hosts"
	$hostsMatches = Select-String -Path $hostsPath -Pattern '^\s{0,}((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))\s{1,}([A-z.\-0-9]{1,})\s{0,}(?:#(\s{0,}[A-z\s]{1,})|\s{0,})$' -AllMatches | % { $_.Matches } 
	Write-LogMessage -Type Info -Msg "`tFound $($hostsMatches.Count) matches"	
	
	$hostsRecords = @()
	Foreach ($match in $hostsMatches)
	{
		if(($match.Value -ne $null) -and (![string]::IsNullOrEmpty($match.Value.Trim())))
		{
			$record = "" | select IP, DNS, Description
			$record.IP = $match.Groups[1]
			$record.DNS = $match.Groups[5].ToString().Trim()
			$record.Description = $match.Groups[6].ToString().Trim()
			
			$hostsRecords += $record
		}
	}
	
	return $hostsRecords
}
#endregion # Helper Functions

# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-LogMessage "Collecting Vault Files" -Type Debug
Write-LogMessage "Collecting logs between $TimeframeFrom to $TimeframeTo" -Type Debug

$arrVaultFilePaths = @()

# Create a file with the relevant file versions
Write-LogMessage "Collecting Vault file versions and additional information" -Type Debug
$vaultVersions = "$DestFolderPath\_VaultFileVersions.txt"
$ldapsCertificateLog = "$DestFolderPath\_CertificatesValidation.log"
"DBMain:"+$(Get-FileVersion "$ComponentPath\dbmain.exe") | Out-File $vaultVersions
"Logic Container:"+$(Get-FileVersion "$ComponentPath\LogicContainer\BLServiceApp.exe") | Out-File $vaultVersions -append
"DNS Service status:"+$(Check-Service "DNS Client") | Out-File $vaultVersions -append
$EPVVersion = Get-Content $DestFolderPath\_VaultFileVersions.txt | Where-Object { $_.Contains("DBMain:") }
$MyVerson = $EPVVersion.Split(":")
$CompVer = $MyVerson[1].Replace(".","").Substring(0,4)
Write-LogMessage "Collecting Vault files list by timeframe" -Type Debug
$arrVaultFilePaths += $vaultVersions
$arrVaultFilePaths += $ldapsCertificateLog
#
# run CAVaultManager.exe to run the DiagnoseDBReport 3x pausing 60 seconds between runs
$DBRCnt=1
$DiagnoseDBReportParms="DiagnoseDBReport"
$CAVaultManager="$ComponentPath\CAVaultManager.exe"
while ($DBRCnt -le 3) {
	$TimeStamp=Get-Date -format "yyyymmddTHHmmss"
	$DBR = & $CAVaultManager $DiagnoseDBReportParms
	$DBR | Out-File -FilePath "$ComponentPath\DiagnoseDBReport_$TimeStamp.txt"
	Start-Sleep -Seconds 60	
	$DBRCnt++
}
$arrVaultFilePaths += (Get-FilePath "$ComponentPath\DiagnoseDBReport_*.txt")
If ($CompVer -ge 1005) {
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\dbparm.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\paragent.ini")
	$arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\passparm.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\tsparm.ini")
	$arrVaultFilePaths += (Get-FilePath "$ComponentPath\Conf\vault.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Database\my.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Database\vaultdb.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\italog.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\trace.d*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\CAVaultManager.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\paragent.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\VaultConfiguration.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\CACert.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\stats.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\InstallMySQLService.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Logs\old\*.*")
    $arrVaultFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\Archive\ARC*.log" -From $TimeframeFrom -To $TimeframeTo)
    $arrVaultFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Logs\Archive\ARC*.LC.log" -From $TimeframeFrom -To $TimeframeTo)
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Event Notification Engine\Conf\*.*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Event Notification Engine\Logs\*.*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Event Notification Engine\Logs\old\*.*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\LogicContainer\Logs\*.*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\LogicContainer\Config\*.*")
} else {
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\dbparm.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\paragent.ini")
	$arrVaultFilePaths += (Get-FilePath "$ComponentPath\passparm.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\tsparm.ini")
	$arrVaultFilePaths += (Get-FilePath "$ComponentPath\vault.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\trace.d*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\CACert.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\CAVaultManager.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\InstallMySQLService.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\italog.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\paragent.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\VaultConfiguration.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\old\*.*")
    $arrVaultFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Archive\ARC*.log" -From $TimeframeFrom -To $TimeframeTo)
    $arrVaultFilePaths += (Get-FilesByTimeframe -Path "$ComponentPath\Archive\ARC*.LC.log" -From $TimeframeFrom -To $TimeframeTo)
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Database\my.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Database\vaultdb.log")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Event Notification Engine\*.ini")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Event Notification Engine\Logs\*.*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\Event Notification Engine\Logs\old\*.*")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\LogicContainer\*.svclog")
    $arrVaultFilePaths += (Get-FilePath "$ComponentPath\LogicContainer\BLServiceApp.exe.config")
}

#region Check for LDAP certificates
Write-LogMessage -Type Info -Msg "Inspecting hosts from Hosts file, searching for LDAPS addresses"
ForEach ($record in $(Get-HostsRecords))
{
	Write-LogMessage -Type Info -Msg "`tChecking $($record.DNS) ($($record.Description)) for LDAPS certificate chain validation"
	Check-CertificateChain -Server $record.DNS -Port 636
}
Write-LogMessage -Type Info -Msg "Finished Inspecting hosts for LDAPS certificates, please see the CertificatesValidation.log for more details"
#endregion


Collect-Files -arrFilesPath $arrVaultFilePaths -destFolder $DestFolderPath
#
# delete the DiagnoseDBReports just generated
Remove-Item -Force "$ComponentPath\DiagnoseDBReport_*.txt"
Write-LogMessage "Done Collecting Vault Files" -Type Debug