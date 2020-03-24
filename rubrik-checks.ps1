<#
#https://rubrik/swagger-ui/
#https://rubrik-cluster-fqdn-or-ip/docs/v1/
# syntax:
#	$credential = Get-Credential -UserName admin -Message "Enter the password to use for this connection"
#	$credential.Password | ConvertFrom-SecureString | Set-Content <file_path_encrypted_password.txt>
#	$rubrikReport = .\rubrik-documentor.ps1 -server <rubrikIP> -Passwordfile <file_path_encrypted_password.txt>
#	$rubrikReport | Export-CliXML -Path .\customerRubrik.xml
#
#
#	Syntax:
# rubrikcreds.txt must contain the password in an encrypted format.
.\rubrik-checks.ps1 -Passwordfile "D:\scripts\rubrik\rubrikcreds.txt" -username "DOMAIN\administrator" -Target "brik01.lan.local"

#>
param (
	[Parameter(Mandatory = $true,Position = 0,HelpMessage = 'Rubrik FQDN or IP address')]
     [ValidateNotNullorEmpty()]
     [String]$Target,
	[Parameter(Mandatory = $true,Position = 0,HelpMessage = 'Rubrik username (default is admin)')]
     [ValidateNotNullorEmpty()]
     [String]$username,
	#[Parameter(Mandatory = $false,Position = 0,HelpMessage = 'File that contains the encrypted password for the username to be used in this connection')]
	 #[ValidateNotNullorEmpty()][String]$Passwordfile,
	[Parameter(Mandatory = $true)][string]$Passwordfile,
	[Parameter(Mandatory = $false)][int]$showLastMonths=3,
	[Parameter(Mandatory = $false)][string]$logDir=".$([IO.Path]::DirectorySeparatorChar)output",
	[Parameter(Mandatory = $false)][bool]$LOCALDEBUG=$false,
	[Parameter(Mandatory = $false)][int]$reportPeriodDays=30,
	[Parameter(Mandatory = $false)][int]$logsCount=100,
	[Parameter(Mandatory = $false)][int]$topItemsOnly=10,
	[Parameter(Mandatory = $false)][int]$headerType=1,
	[Parameter(Mandatory = $false)][string]$htmlFile="rurikreport.html",
	[Parameter(Mandatory = $false)][string]$reportIntro="The results of a Rubrik health check can be found below.",
	[Parameter(Mandatory = $false)][string]$reportHeader="Rubrik Health Check Report",
	[Parameter(Mandatory = $true)][string]$farmName,
	[Parameter(Mandatory = $true)][string]$itoContactName
)

#$silencer = Import-Module -Name ".\rubrik-Module.psm1" -Force:$true
$silencer = Import-Module -Name ".\genericModule.psm1" -Force:$true
Set-Variable -Name scriptName -Value $($MyInvocation.MyCommand.name) -Scope Global
set-variable -Name logDir -Scope global -Value $logDir
if (!(Get-Variable -Name logDir -Scope Global -ErrorAction SilentlyContinue)) {
	set-variable -Name logDir -Scope global -Value $logDir
}

logThis -msg "Running $($global:scriptName)"
#Set-Variable -Name logDir -Value $logDir -Scope Global
#Set-Variable -Name reportIndex -Value "$logDir\index.txt" -Scope Global
$type="Rubrik"
$global:results=@{}
$global:results["$type"]=@{}
$global:results["$type"]["Reports"]=@{}
$global:results["$type"]["Reports"]["Capacity"]=@{}
$global:results["$type"]["Collection"] = @{}
#Write-Host ":: Log File $global:logfile"

$MyCredentials = Get-myCredentials -SecureFileLocation $Passwordfile -User $username

$connResults = Connect-Rubrik -Server $Target -Username $username -Password $MyCredentials.password
$session = transposeNameValueTableToArray -keyValueTable $connResults

#schema xml.VeeamBackup.Reports.Daily Checks.MetaData
#schema xml.VeeamBackup.Reports.Daily Checks.DataTable
$xml=@{}
$reportCategory="Rubrik"
$xml["$reportCategory"]=@{}
$xml.$reportCategory["Runtime"]=@{}
$xml.$reportCategory.Runtime["Name"] = $session.Server
$xml.$reportCategory.Runtime["Report Ran on"]=$session.time
$xml.$reportCategory.Runtime["Session Information"]=$session
$xml.$reportCategory.Runtime["Type"] = "Health Checks"
$xml.$reportCategory.Runtime["Reporting Period (Months)"] = $showLastMonths
$xml.$reportCategory["Infrastructure"]=@{} # where all of the collection will go
$xml.$reportCategory["Reports"]=@{} # where all of the reports will be go
$xml.$reportCategory.Reports["Capacity"]=@{} # where all of the reports will be go
$reportIndex=-1


#Collection
# SHOW List of TSM Servers audit in this health check.
$showClusterSummary=$true
if ($showClusterSummary)
{
	$reportIndex++
	$title="Rubrik Cluster Information"

	$tableHeader="$title"
	Write-Host "Processing $tableHeader Report.."
	$metaInfo = @()
	$metaInfo +="tableHeader=$title"
	$metaInfo +="introduction=Table 'x' provides a summary of the Rubrik appliance."
	$metaInfo +="chartable=false"
	$metaInfo +="titleHeaderType=h$($headerType)"
	#$metaInfo +="calculatetotals=1,2,3,4,5,6,7,8,9,10" # Calculate Total for columns 1,2,3,xxxx - Note, column 0 is the first column.
	$metaInfo +="displayTableOrientation=List" # options are List or Table
	$xml.$reportCategory.Reports.Capacity[$reportIndex] = @{}
	$clusterInfo = Get-RubrikClusterInfo
	$clusterStorage = Get-RubrikClusterStorage

	$dataTable = New-Object System.Object
	$dataTable | Add-Member -MemberType NoteProperty -Name "Cluster Name" -Value $clusterInfo.Name
	$dataTable | Add-Member -MemberType NoteProperty -Name "Platform" -Value $clusterInfo.Platform
	$dataTable | Add-Member -MemberType NoteProperty -Name "Software Version" -Value $clusterInfo.softwareVersion
	$dataTable | Add-Member -MemberType NoteProperty -Name "Status" -Value $clusterInfo.ClusterStatus
	$dataTable | Add-Member -MemberType NoteProperty -Name "Brik Count" -Value $clusterInfo.BrikCount
	$dataTable | Add-Member -MemberType NoteProperty -Name "Node Count" -Value $clusterInfo.NodeCount
	$dataTable | Add-Member -MemberType NoteProperty -Name "Integrated with Polaris" -Value $clusterInfo.ConnectedToPolaris
	$dataTable | Add-Member -MemberType NoteProperty -Name "Capacity" -Value "$($clusterInfo.CPUCoresCount) CPU Cores, $($clusterInfo.MemoryCapacityInGb) GB of Memory, $($clusterStorage.DiskCapacityInTb) TB of Capacity"

	$clusterStorage.GetEnumerator() | Select-Object Key,Value | ForEach-Object {
		$clusterStorageKey = $_.Key
		$clusterStorageValue = $_.Value
		$string = ($clusterStorageKey.substring(0,1).toupper() + $clusterStorageKey.substring(1) -creplace '[A-Z]', ' $&').Trim();
		$dataTable | Add-Member -MemberType NoteProperty -Name "$string" -Value $clusterStorageValue
	}
	$xml.$reportCategory.Reports.Capacity.$reportIndex["MetaData"] = $metaInfo
	$xml.$reportCategory.Reports.Capacity.$reportIndex["DataTable"] = $dataTable
	# Show
	#$xml.$reportCategory.Reports.Capacity.$reportIndex.DataTable

}


# List Cluster Nodes
$showClusterNodes=$true
if ($showClusterNodes)
{
	$reportIndex++
	$title="Rubrik Nodes"

	$tableHeader="$title"
	logThis -msg "Processing $tableHeader Report"
	$metaInfo = @()
	$metaInfo +="tableHeader=$title"
	$metaInfo +="introduction=Table 'x' presents a list of server components of the Rubrik appliance."
	$metaInfo +="chartable=false"
	$metaInfo +="titleHeaderType=h$($headerType)"
	#$metaInfo +="calculatetotals=1,2,3,4,5,6,7,8,9,10" # Calculate Total for columns 1,2,3,xxxx - Note, column 0 is the first column.
	$metaInfo +="displayTableOrientation=Table" # options are List or Table
	$xml.$reportCategory.Reports.Capacity[$reportIndex] = @{}
	$clusterNodes = Get-RubrikNode
	$dataTable = @()
	$clusterNodes | sort-object -Property BrikId | ForEach-Object {
		$clusterNode = $_
		$obj = New-Object System.Object
		$obj | Add-Member -MemberType NoteProperty -Name "Node Id" -Value $clusterNode.Id
		$obj | Add-Member -MemberType NoteProperty -Name "Brik Id" -Value $clusterNode.BrikId
		$obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $clusterNode.ipAddress
		$obj | Add-Member -MemberType NoteProperty -Name "Status" -Value $clusterNode.Status
		$dataTable += $obj
	}
	$xml.$reportCategory.Reports.Capacity.$reportIndex["MetaData"] = $metaInfo
	$xml.$reportCategory.Reports.Capacity.$reportIndex["DataTable"] = $dataTable
}

# List SLAs
$showSLAs=$true
if ($showSLAs)
{
	$reportIndex++
	$title="Backup SLA"

	$tableHeader="$title"
	Write-Host "Processing $tableHeader Report"
	$xml.$reportCategory.Reports.Capacity[$reportIndex] = @{}
	$metaInfo = @()
	$metaInfo +="tableHeader=$title"
	$metaInfo +="introduction=Table 'x' provides a summary of configured Rubrik SLAs."
	$metaInfo +="chartable=false"
	$metaInfo +="titleHeaderType=h$($headerType)"
	#$metaInfo +="calculatetotals=1,2,3,4,5,6,7,8,9,10" # Calculate Total for columns 1,2,3,xxxx - Note, column 0 is the first column.
	$metaInfo +="displayTableOrientation=Table" # options are List or Table
	$clusterSLAs = Get-RubrikSLA

	$dataTable = @()
	$clusterSLAs | sort-object -Property Name | ForEach-Object {
		Write-Host "`t-> Processing SLA :- $($clusterSLA.Name)"
		$clusterSLA = $_
		$obj = New-Object System.Object
		if ($clusterSLA.isDefault)
		{
			$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value "$($clusterSLA.Name) (default)"
		} else {
			$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value "$($clusterSLA.Name)"
		}
		# Work out frequencies
		$backupSchedule = ""
		$frequenciesList = ($clusterSLA | Select-Object -ExpandProperty Frequencies | Get-Member -MemberType NoteProperty).Name | Where-Object {$_ -ne "Name"}
		$frequenciesList | ForEach-Object {
			$scheduleName=$_
			$retentionReword = $scheduleName -replace 'ly','s' -replace 'dais','days'
			$retention ="to keep for $($clusterSLA.frequencies.$scheduleName.retention) $retentionReword"
			$backupSchedule += "$($clusterSLA.frequencies.$scheduleName.frequency) x $scheduleName backups $retention,`n"
		}
		$obj | Add-Member -MemberType NoteProperty -Name "Schedules and Retention" -Value "$backupSchedule"

		# Iterate through the protected objects
		$objectsProtectedString="$($clusterSLA.'numProtectedObjects') in Total"
		$objectsProtected = ($clusterSLA | Get-Member -MemberType NoteProperty | Where-Object {$_.Name -like "num*"}).Name | Where-Object {$_ -ne "numProtectedObjects"}
		$objectsProtected | ForEach-Object {
			$objectName = $_
			if ($clusterSLA.$objectName -gt 0)
			{
				$objectNameShorthand = $objectName -replace 'num',''
				$objectsProtectedString += "`n$($clusterSLA.$objectName) x $objectNameShorthand"
			}
		}

		$obj | Add-Member -MemberType NoteProperty -Name "Protected Objects" -Value "$objectsProtectedString"

		$dataTable += $obj
	}
	$xml.$reportCategory.Reports.Capacity.$reportIndex["MetaData"] = $metaInfo
	$xml.$reportCategory.Reports.Capacity.$reportIndex["DataTable"] = $dataTable
}

# List Client being backed up
$showClientVMs=$true
if ($showClientVMs)
{
	# Gather the information required
	#$vms = Get-RubrikVM | where-Object {$_.Name -notlike "*OmniStackVC*"} | Sort-Object Name -Unique | Select-Object -First 5
	$vms = Get-RubrikVM | where-Object {$_.Name -notlike "*OmniStackVC*"} | Sort-Object Name -Unique
	$totalVMs = ($vms | measure-object).Count
	$unprotectedVMCount = ($vms | Where-Object {$_.effectiveSlaDomainName -like "Unprotected"} | Measure-Object).Count
	$protectedVMs = $vms | Where-Object {$_.effectiveSlaDomainName -notlike "Unprotected"}
	$protectedVMsCount = ($protectedVMs | Measure-Object).Count


	#########
	#
	# VM Summary
	$reportIndex++
	$title="VM Protection Summary"
	$tableHeader="$title"
	Write-Host "Processing $tableHeader Report.."
	$metaInfo = @()
	$metaInfo +="tableHeader=$title"
	$metaInfo +="introduction=Table 'x' provides a summary of VM backups."
	$metaInfo +="chartable=false"
	$metaInfo +="titleHeaderType=h$($headerType)"
	#$metaInfo +="calculatetotals=1,2,3,4,5,6,7,8,9,10" # Calculate Total for columns 1,2,3,xxxx - Note, column 0 is the first column.
	$metaInfo +="displayTableOrientation=List" # options are List or Table
	$xml.$reportCategory.Reports.Capacity[$reportIndex] = @{}
	$dataTable = @()
	$obj = New-Object System.Object
	$obj | Add-Member -MemberType NoteProperty -Name "Virtual Machines" -Value "$totalVMs x Registered, $protectedVMsCount x Protected, $unprotectedVMCount x Unprotected"

	# List used SLAs
	$perSLACountString=""
	$protectedVMs | Group-Object -Property effectiveSlaDomainName | Select-Object Name,Count | ForEach-Object {
		$groupBySLADomain = $_
		$perSLACountString += "$($groupBySLADomain.Count) x $($groupBySLADomain.Name), "
	}
	$obj | Add-Member -MemberType NoteProperty -Name "SLAs Usage" -Value "$perSLACountString"

	# List Unused SLAs
	$usedSLAsName = ($protectedVMs | Group-Object -Property effectiveSlaDomainName | Select-Object Name).Name | Sort-Object
	$unusedSLANamesString=""
	(Get-RubrikSLA | select-Object Name).Name | sort-Object | Where-Object {$usedSLAsName -notcontains $_ } | Where-Object {
		$unusedSLANamesString += "$_, "
	}

	$obj | Add-Member -MemberType NoteProperty -Name "Unused SLAs" -Value "$unusedSLANamesString"
	$dataTable = $obj
	$xml.$reportCategory.Reports.Capacity.$reportIndex["MetaData"] = $metaInfo
	$xml.$reportCategory.Reports.Capacity.$reportIndex["DataTable"] = $dataTable
	#$xml.Rubrik.Reports.$title

	############## LIST VMS
	$reportIndex++
	$title="Protected Virtual Machines"

	$tableHeader="$title"
	Write-Host "Processing $tableHeader Report.."
	$metaInfo = @()
	$metaInfo +="tableHeader=$title"
	$metaInfo +="introduction=Table 'x' provides a summary of VMs."
	$metaInfo +="chartable=false"
	$metaInfo +="titleHeaderType=h$($headerType)"
	#$metaInfo +="calculatetotals=1,2,3,4,5,6,7,8,9,10" # Calculate Total for columns 1,2,3,xxxx - Note, column 0 is the first column.
	$metaInfo +="displayTableOrientation=Table" # options are List or Table
	$xml.$reportCategory.Reports.Capacity[$reportIndex] = @{}
	$vmindex=1
	$dataTable = @()
	$protectedVMs | sort-object -Property Name  | ForEach-Object {
		$protectedVM = $_
		Write-Host "`t-> Processing VM $vmindex/$(($protectedVMs | measure-object).Count):- $($protectedVM.Name.ToUpper()).."
		$vmEvents = Get-RubrikEvent -Objectname $protectedVM.Name -Status Failure -Limit 1000
		#$recentBackup = $events | Sort-Object -Property Date -Descending | select-object -First 1
		#$oldestBackup = $events | Sort-Object -Property Date -Descending | select-object -Last 1
		$snapshotBackups = $protectedVM | Get-RubrikSnapshot | Sort-Object Date
		$recentBackup = $snapshotBackups | Select-Object -Last 1
		$oldestBackup  = $snapshotBackups | Select-Object -First 1
		$obj = New-Object System.Object
		$obj | Add-Member -MemberType NoteProperty -Name "VM Name" -Value "$($protectedVM.Name.ToUpper())"
		$obj | Add-Member -MemberType NoteProperty -Name "SLA" -Value "$($protectedVM.effectiveSlaDomainName)"
		$obj | Add-Member -MemberType NoteProperty -Name "Cluster" -Value "$($protectedVM.clusterName)"

		if ( (($vmEvents | measure-object).Count -gt 0) -and ($vmEvents.data))
		{
			#Write-Host $vmEvents
			#$vmEvents | Export-CliXML -Path "$($global:logDir)\$($protectedVM.Name.ToUpper())_failures_all.xml"
			#pause
			# if the event is less than 30 days process it
			$eventsString="None"
			$eventErrorCount=0
			$eventsLast30days = @()
			$vmEvents | ForEach-Object {
				$myEvent = $_
				$eventsString=""
				if ((get-date $myEvent.date) -gt (get-date).AddDays(-30))
				{
					$eventErrorCount++
					$eventsLast30days += $myEvent
					#$eventInfo=$myEvent.eventInfo -replace '{"message":','' -replace '"','' -replace ',"params":{}}',''
					#$eventsString += "$eventInfo"
				}
			}
			#$eventErrorCount++
			#if ($eventErrorCount -gt 1)
			#{
				#$eventsLast30days | Export-CliXML -Path "$($global:logDir)\$($protectedVM.Name.ToUpper())_failures_30DaysLess.xml"
			#}
			$obj | Add-Member -MemberType NoteProperty -Name "Failures (last 30 days)" -Value "$eventErrorCount errors found, $eventsString"
		} else {
			$obj | Add-Member -MemberType NoteProperty -Name "Failures (last 30 days)" -Value "None"
		}
		#$obj | Add-Member -MemberType NoteProperty -Name "Oldest Backup" -Value "$($oldestBackup.Date) ($($oldestBackup.eventStatus))"
		<#
		id                     : 1f952864-7a45-44ad-9f8a-eb1920e968c5
slaName                : General-Daily
slaId                  : e2c48b5e-bf20-4989-9b53-29ba134d7404
replicationLocationIds :
cloudState             : 6
indexState             : 1
isOnDemandSnapshot     : False
archivalLocationIds    : 3b441e9f-e1c5-47eb-8ef7-2c02efe325b3
date                   : 2019-11-20T07:15:08.596Z
vmName                 : VCSWKSTEST04
consistencyLevel       : CRASH_CONSISTENT

#>
		$backupString = "$(($snapshotBackups | measure-object).Count) backups found"
		$backupString += ", Latest taken on $((get-date $recentBackup.date))"
		$backupString += ", Oldest taken on $((get-date $oldestBackup.date))"
		$snapshotBackups | Group-Object cloudState | Select-Object Name,Count | ForEach-Object {
			$backupGroup = $_
			switch ($backupGroup.Name)
			{
				0 { $backupString += ", $($backupGroup.Count) stored locally"}
				2 { $backupString += ", $($backupGroup.Count) archived to azure"}
				6 { $backupString += ", $($backupGroup.Count) archived locally and to Azure"}
				default { }
			}
		}

		$obj | Add-Member -MemberType NoteProperty -Name "Backups" -Value "$backupString"
		$dataTable += $obj
		$vmindex++
	}

	$xml.$reportCategory.Reports.Capacity.$reportIndex["MetaData"] = $metaInfo
	$xml.$reportCategory.Reports.Capacity.$reportIndex["DataTable"] = $dataTable
}

$showFailures=$false
if ($showFailures)
{
	$reportIndex++
	$failures = Get-RubrikEvent -Status "Failure" -ShowOnlyLatest
	$title="Recent Events"
	$tableHeader="$title"
	Write-Host "Processing $tableHeader Report.."
	$metaInfo = @()
	$metaInfo +="tableHeader=$title"
	$metaInfo +="introduction=Table 'x' provides a summary recent failure events. A total of $(($failures | Measure-Object).Count) were found."
	$metaInfo +="chartable=false"
	$metaInfo +="titleHeaderType=h$($headerType)"
	#$metaInfo +="calculatetotals=1,2,3,4,5,6,7,8,9,10" # Calculate Total for columns 1,2,3,xxxx - Note, column 0 is the first column.
	$metaInfo +="displayTableOrientation=List" # options are List or Table
	$xml.$reportCategory.Reports.Capacity[$reportIndex] = @{}
	$dataTable = @()

	$failures | group-object objectType | ForEach-Object {
		$eventObjectType = $_
		$eventObjectType.Group | Group-Object EventType | ForEach-Object {
			$event = $_
			$message = @()
			$event.Group.eventinfo  | ForEach-Object {
				$message += $_ -replace '{"message":','' -replace '"','' -replace ',params:{}}',''
				$obj = New-Object System.Object
				$obj | Add-Member -MemberType NoteProperty -Name "Type" -Value "$($eventObjectType.Name)" #$($event.Name)"
				$obj | Add-Member -MemberType NoteProperty -Name "Message" -Value "$message"
				$dataTable += $obj
			}

		}
	}

	$xml.$reportCategory.Reports.Capacity.$reportIndex["MetaData"] = $metaInfo
	$xml.$reportCategory.Reports.Capacity.$reportIndex["DataTable"] = $dataTable
}

$generateHTMLReport=$true
if ($generateHTMLReport)
{
	$htmlPage = generateHTMLReport -xml $xml -reportHeader $reportHeader -reportIntro $reportIntro -farmName $farmName -itoContactName $itoContactName
	$htmlPage | Out-File "$($global:logDir)$($([IO.Path]::DirectorySeparatorChar))$htmlFile"
	logThis -msg "---> Opening $htmlFile"
	#if ($global:report.Runtime.Configs.openReportOnCompletion)
	#{

		& "$($global:logDir)$($([IO.Path]::DirectorySeparatorChar))$htmlFile"
	#}
}
return $xml