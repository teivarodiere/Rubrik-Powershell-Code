#rubrik-Module.psm1
# Download powershgell modules from https://github.com/rubrikinc/rubrik-sdk-for-powershell
$silencer = Import-Module -Name "..$([IO.Path]::DirectorySeparatorChar)..$([IO.Path]::DirectorySeparatorChar)genericModule.psm1" -Force:$true

<#
# Check for pre-requisites module
$preReqModules="Rubrik"
$preReqModules | ForeEach-object {
    if (Get-Module -ListAvailable -Name S_) {
        Write-Host "Module exists"
    }
    else {
        Write-Host "Module does not exist"
        $silencer = Import-Module Rubrik -Global -Force
        exit 1
    }
}

#Requires -Version 3

function createResultsObject()
{
    $global:results=@{}
}

function createResultsObjectSubType([String] $type)
{
    if ($global:results)
    {
        $global:results["$type"]=@{}
    } else {
        createResultsObject()
        createResultsObjectSubType -type $type
    }
}
function createResultsObjectSubTypeReport([String] $type)
{
    if ($global:results.$type)
    {        
        $global:results["$type"]["Reports"]=@{}
    } else {
        createResultsObjectSubType -type $type
        createResultsObjectSubTypeReport -type $type
    }
}

function createResultsObjectSubTypeReportCapacity([String] $type)
{
    if ($global:results.$type.$report)
    {        
        $global:results["$type"]["Reports"]=@{}
    } else {
        createResultsObjectSubTypeReport -type $type
        createResultsObjectSubTypeReportCapacity -type $type
    }
}


$global:results["$type"]["Collection"] = @{}

#>


<# OLD CODE
#https://rubrik/swagger-ui/
param ( 
	[Parameter(Mandatory = $true,Position = 0,HelpMessage = 'Rubrik FQDN or IP address')]
     [ValidateNotNullorEmpty()]
     [String]$server,
	[Parameter(Mandatory = $false,Position = 0,HelpMessage = 'Rubrik username')]
     [ValidateNotNullorEmpty()]
     [String]$username="admin",
	[Parameter(Mandatory = $true,Position = 0,HelpMessage = 'File that contains the encrypted password for the username to be used in this connection')]
     [ValidateNotNullorEmpty()]
     [String]$SecureFileLocation
)

Import-Module -Force -Name ".\rubrik-Module.psm1"

$session = connect -Server $server -Username $username -SecureFileLocation $SecureFileLocation

$get_requests= @('host','vcenter','datacenter','oracledb','clusterIps','compute_cluster','slaDomain','mount','system/version','vm','vm/list','virtual/disk','system/ntp/servers','support/tunnel') #,'internal/job/type/backup','internal/config/crystal','report/vm')
$report = @{}
$get_requests | %{
	$request=$_
	Write-Host "Processing $request"
	#$fieldname = $request -split '/' | select -Last 1
	$fieldname = $request -replace '/','_' #| select -Last 1
	try {	
		$response = Invoke-WebRequest -Uri "$($session.ConnectionUri)/$request" -Headers $session.Headers -Method Get
		$jsonResponse = ConvertFrom-Json -InputObject $response.Content
		$report[$fieldname] = $jsonResponse
	} catch 
	{
		throw "Error gathering ""$request"" from $($session.Servername)"
	}
}

return $report


#>

<##

Rurbrik has provided a list of commands with it's new Rubrik powershell module.

Key                                Value
---                                -----
Connect-Rubrik                     Connect-Rubrik
Disconnect-Rubrik                  Disconnect-Rubrik
Export-RubrikDatabase              Export-RubrikDatabase
Export-RubrikReport                Export-RubrikReport
Get-RubrikAPIVersion               Get-RubrikAPIVersion
Get-RubrikAvailabilityGroup        Get-RubrikAvailabilityGroup
Get-RubrikDatabase                 Get-RubrikDatabase
Get-RubrikDatabaseFiles            Get-RubrikDatabaseFiles
Get-RubrikDatabaseMount            Get-RubrikDatabaseMount
Get-RubrikDatabaseRecoverableRange Get-RubrikDatabaseRecoverableRange
Get-RubrikFileset                  Get-RubrikFileset
Get-RubrikFilesetTemplate          Get-RubrikFilesetTemplate
Get-RubrikHost                     Get-RubrikHost
Get-RubrikHyperVVM                 Get-RubrikHyperVVM
Get-RubrikLogShipping              Get-RubrikLogShipping
Get-RubrikManagedVolume            Get-RubrikManagedVolume
Get-RubrikManagedVolumeExport      Get-RubrikManagedVolumeExport
Get-RubrikMount                    Get-RubrikMount
Get-RubrikNASShare                 Get-RubrikNASShare
Get-RubrikNutanixVM                Get-RubrikNutanixVM
Get-RubrikOrganization             Get-RubrikOrganization
Get-RubrikReport                   Get-RubrikReport
Get-RubrikReportData               Get-RubrikReportData
Get-RubrikRequest                  Get-RubrikRequest
Get-RubrikSLA                      Get-RubrikSLA
Get-RubrikSnapshot                 Get-RubrikSnapshot
Get-RubrikSoftwareVersion          Get-RubrikSoftwareVersion
Get-RubrikSQLInstance              Get-RubrikSQLInstance$vm
Get-RubrikSupportTunnel            Get-RubrikSupportTunnel
Get-RubrikUnmanagedObject          Get-RubrikUnmanagedObject
Get-RubrikVersion                  Get-RubrikVersion
Get-RubrikVM                       Get-RubrikVM
Get-RubrikVMSnapshot               Get-RubrikVMSnapshot
Get-RubrikVolumeGroup              Get-RubrikVolumeGroup
Get-RubrikVolumeGroupMount         Get-RubrikVolumeGroupMount


Invoke-RubrikRESTCall              Invoke-RubrikRESTCall
Move-RubrikMountVMDK               Move-RubrikMountVMDK
New-RubrikDatabaseMount            New-RubrikDatabaseMount
New-RubrikFileset                  New-RubrikFileset
New-RubrikFilesetTemplate          New-RubrikFilesetTemplate
New-RubrikHost                     New-RubrikHost
New-RubrikLogBackup                New-RubrikLogBackup
New-RubrikLogShipping              New-RubrikLogShipping
New-RubrikManagedVolume            New-RubrikManagedVolume
New-RubrikManagedVolumeExport      New-RubrikManagedVolumeExport
New-RubrikMount                    New-RubrikMount
New-RubrikNASShare                 New-RubrikNASShare
New-RubrikReport                   New-RubrikReport
New-RubrikSLA                      New-RubrikSLA
New-RubrikSnapshot                 New-RubrikSnapshot
New-RubrikVMDKMount                New-RubrikVMDKMount
New-RubrikVolumeGroupMount         New-RubrikVolumeGroupMount
Protect-RubrikDatabase             Protect-RubrikDatabase
Protect-RubrikFileset              Protect-RubrikFileset
Protect-RubrikHyperVVM             Protect-RubrikHyperVVM
Protect-RubrikNutanixVM            Protect-RubrikNutanixVM
Protect-RubrikTag                  Protect-RubrikTag
Protect-RubrikVM                   Protect-RubrikVM
Remove-RubrikDatabaseMount         Remove-RubrikDatabaseMount
Remove-RubrikFileset               Remove-RubrikFileset
Remove-RubrikHost                  Remove-RubrikHost
Remove-RubrikLogShipping           Remove-RubrikLogShipping
Remove-RubrikManagedVolume         Remove-RubrikManagedVolume
Remove-RubrikManagedVolumeExport   Remove-RubrikManagedVolumeExport
Remove-RubrikMount                 Remove-RubrikMount
Remove-RubrikNASShare              Remove-RubrikNASShare
Remove-RubrikReport                Remove-RubrikReport
Remove-RubrikSLA                   Remove-RubrikSLA
Remove-RubrikUnmanagedObject       Remove-RubrikUnmanagedObject
Remove-RubrikVolumeGroupMount      Remove-RubrikVolumeGroupMount
Reset-RubrikLogShipping            Reset-RubrikLogShipping
Restore-RubrikDatabase             Restore-RubrikDatabase
Set-RubrikAvailabilityGroup        Set-RubrikAvailabilityGroup
Set-RubrikBlackout                 Set-RubrikBlackout
Set-RubrikDatabase                 Set-RubrikDatabase
Set-RubrikHyperVVM                 Set-RubrikHyperVVM
Set-RubrikLogShipping              Set-RubrikLogShipping
Set-RubrikManagedVolume            Set-RubrikManagedVolume
Set-RubrikMount                    Set-RubrikMount
Set-RubrikNASShare                 Set-RubrikNASShare
Set-RubrikNutanixVM                Set-RubrikNutanixVM
Set-RubrikSQLInstance              Set-RubrikSQLInstance
Set-RubrikSupportTunnel            Set-RubrikSupportTunnel
Set-RubrikVM                       Set-RubrikVM
Start-RubrikManagedVolumeSnapshot  Start-RubrikManagedVolumeSnapshot
Stop-RubrikManagedVolumeSnapshot   Stop-RubrikManagedVolumeSnapshot
Sync-RubrikAnnotation              Sync-RubrikAnnotation
Sync-RubrikTag                     Sync-RubrikTag
#>

<# OLD COODE
#rubrik-Module.psm1
#
function connect ( 
	[Parameter(Mandatory = $true,Position = 0,HelpMessage = 'Rubrik FQDN or IP address')]
     [ValidateNotNullorEmpty()]
     [String]$server,
	[Parameter(Mandatory = $false,Position = 0,HelpMessage = 'Rubrik username')]
     [ValidateNotNullorEmpty()]
     [String]$username="admin",
	[Parameter(Mandatory = $true,Position = 0,HelpMessage = 'File that contains the encrypted password for the username to be used in this connection')]
     [ValidateNotNullorEmpty()]
     [String]$SecureFileLocation

)
{
	#Import-Module -Force -Name "Rubrik.psd1"
	# Allow untrusted SSL certs
     Add-Type -TypeDefinition @"
	    using System.Net;
	    using System.Security.Cryptography.X509Certificates;
	    public class TrustAllCertsPolicy : ICertificatePolicy {
	        public bool CheckValidationResult(
	            ServicePoint srvPoint, X509Certificate certificate,
	            WebRequest request, int certificateProblem) {
	            return true;
	        }
	    }
"@
     [System.Net.ServicePointManager]::CertificatePolicy = New-Object -TypeName TrustAllCertsPolicy
	# Create secure credentials to use in the remaining script.
	
	
	if ($SecureFileLocation)
	{	 
		$password = Get-Content $SecureFileLocation | ConvertTo-SecureString 
		$credential = New-Object System.Management.Automation.PsCredential($username,$password)

	}
	if (-not $credential)
	{
		$credential = Get-Credential -UserName $username -Message "Enter the password to use for this connection"
		#$password = $credentialObject.Password | ConvertFrom-SecureString 
	}

	
	
	try 
	{				
		$rubrik = Connect-Rubrik -Server $server -Username $username -Password $credential.Password
		return $rubrik
	}
	catch 
	{
		throw "Error connecting to Rubrik server ""$server"""
		return $null
	}
}


#Requires -Version 3
function Connect-Rubrik 
{
    #<#  
            .SYNOPSIS
            Connects to Rubrik and retrieves a token value for authentication
            .DESCRIPTION
            The Connect-Rubrik function is used to connect to the Rubrik RESTful API and supply credentials to the /login method. Rubrik then returns a unique token to represent the user's credentials for subsequent calls. Acquire a token before running other Rubrik cmdlets.
            .NOTES
            Written by Chris Wahl for community usage
            Twitter: @ChrisWahl
            GitHub: chriswahl
            .LINK
            https://github.com/rubrikinc/PowerShell-Module
            .EXAMPLE
            Connect-Rubrik -Server 192.168.1.1 -Username admin
            This will connect to Rubrik with a username of "admin" to the IP address 192.168.1.1. The prompt will request a secure password.
            .EXAMPLE
            Connect-Rubrik -Server 192.168.1.1 -Username admin -Password (ConvertTo-SecureString "secret" -asplaintext -force)
            If you need to pass the password value in the cmdlet directly, use the ConvertTo-SecureString function.
    ##

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,Position = 0,HelpMessage = 'Rubrik FQDN or IP address')]
        [ValidateNotNullorEmpty()]
        [String]$Server,
        [Parameter(Mandatory = $true,Position = 1,HelpMessage = 'Rubrik username')]
        [ValidateNotNullorEmpty()]
        [String]$Username,
        [Parameter(Mandatory = $true,Position = 2,HelpMessage = 'Rubrik password')]
        [ValidateNotNullorEmpty()]
        [SecureString]$Password,
	   [Parameter(Mandatory = $false,Position = 2,HelpMessage = 'Set this to false if you want to simply return the connection settings instead setting a global variable')]
        [ValidateNotNullorEmpty()]
	   $setglobalVar=$true,
	   [Parameter(Mandatory = $false,Position = 2,HelpMessage = 'Verbose the connection to troubleshoot')]
        [ValidateNotNullorEmpty()]
	   $verboseInfo=$false

    )

    Process {

        # Allow untrusted SSL certs
        Add-Type -TypeDefinition @"
	    using System.Net;
	    using System.Security.Cryptography.X509Certificates;
	    public class TrustAllCertsPolicy : ICertificatePolicy {
	        public bool CheckValidationResult(
	            ServicePoint srvPoint, X509Certificate certificate,
	            WebRequest request, int certificateProblem) {
	            return true;
	        }
	    }
"@
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object -TypeName TrustAllCertsPolicy

        # Build the URI
        $uri = 'https://'+$server+':443/login'

        # Build the login call JSON
        $credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $Password        
        $body = @{
            userId   = $username
            password = $credentials.GetNetworkCredential().Password
        }

        # Submit the token request
        try 
        {
            $r = Invoke-WebRequest -Uri $uri -Method: Post -Body (ConvertTo-Json -InputObject $body)
        }
        catch 
        {
            throw 'Error connecting to Rubrik server'
        }
        $RubrikServer = $server
        $RubrikToken = (ConvertFrom-Json -InputObject $r.Content).token
        if($verboseInfo) { Write-Host -Object "Acquired token: $RubrikToken`r`nYou are now connected to the Rubrik API." }

        # Validate token and build Base64 Auth string
        $auth = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($RubrikToken+':'))
        if ($setglobalVar)
	   {
	   	$global:RubrikServer = $server
	   	$global:RubrikToken = $RubrikToken
	   	$global:RubrikHead = @{
          	  'Authorization' = "Basic $auth"
        	}
	   }

	  return @{
	  	'Servername' = $server
		'Token' = (ConvertFrom-Json -InputObject $r.Content).token
		'Headers' = @{
			'Authorization' = "Basic $auth"
		}
		'ConnectionUri' = "https://$($server):443"
	  }
		

    } # End of process
} # End of function
#>