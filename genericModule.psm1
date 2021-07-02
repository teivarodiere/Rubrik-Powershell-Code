############# CUSTOMER PROPERTIES DECLARATION
<#Add-Type @"
Class MetaInfo
{
	[string]"TableHeader"
	[string]"Introduction"
	[string]"Chartable"
	[int[]]"TitleHeaderType"
	[string]"ShowTableCaption"
	#[string[]]"DisplayTableOrientation"="List","Table","None";
}
"@#>
<#
# You need to define each item as a @(ForgroundColor,BackgroundColor)
# available colors
#    Black
#    Blue
#    Cyan
#    DarkBlue
#    DarkCyan
#    DarkGray
# 	DarkGreen
#    DarkMagenta
#    DarkRed
#    DarkYellow
#    Gray
#    Green
#    Magenta
#    Red
#    White
#    Yellow
#>
$PATHSEPARATOR=[IO.Path]::DirectorySeparatorChar

$global:colours = @{
	"Information"=@{
		"Foreground" = "Yellow"
		"Background" = "Black"
	};
	"Error"=@{
		"Foreground" = "White"
		"Background" = "Red"
	};
	"Change"=@{
		"Foreground" = "Blue"
		"Background" = "Black"
	};
	"NoChange"=@{
		"Foreground" = "Green"
		"Background" = "Black"
	};
	"Highlight"=@{
		"Foreground" = "Cyan"
		"Background" = "Black"
	};

	"Note"=@{
		"Foreground" = "Green"
		"background" = "Black"
	};
	"Alert"=@{
		"Foreground" = "Yellow"
		"background" = "Red"
	};
}

$global:defaultAttributeName="Name"
$global:defaultAttributeId="Id"
$global:defaultAttributeKeys="key"
<#
$global:colours = @{
	"Information"="Yellow";
	"Error"="Red";
	"ChangeMade"="Blue";
	"Highlight"="Green";
	"Note"=""
}
#>
## LOGING OPTIONS
<#
if (!(Get-Variable logTofile)) { [bool]$global:logTofile = $false }
if (!(Get-Variable logInMemory)) { [bool]$global:logInMemory = $true }
if (!(Get-Variable logToScreen)) { [bool]$global:logToScreen = $true }
#>

if (Get-Variable logTofile -Scope Global -ErrorAction Ignore) { } else { New-Variable -Scope Global -Name logTofile -Value $false }

if (Get-Variable logInMemory -Scope Global -ErrorAction Ignore) { } else { New-Variable -Scope Global -Name logInMemory -Value $true }
if (Get-Variable logToScreen -Scope Global -ErrorAction Ignore) { } else { New-Variable -Scope Global -Name logToScreen -Value $true }

function getEventsReportingWindow (
		[Parameter(Mandatory=$true)][int]$lastMonths,
		[Parameter(Mandatory=$false)][int]$lastDayOfReportOveride)
{
	if ($lastDayOfReportOveride)
	{
		$eventsEndPeriod=Get-Date $lastDayOfReportOveride
	} else {
		$eventsEndPeriod=Get-Date
	}
	$eventsStartPeriod = $eventsEndPeriod.AddMOnths(-$lastMonths)
	return ($eventsStartPeriod,$eventsEndPeriod)
}

# this function returns meta data for printing report even thought there are no datatable returned from a query
function returnEmptyMetaData([Parameter(Mandatory=$true)][string]$title, [Parameter(Mandatory=$false)][int]$headerType=1)
{
	$objMetaInfo = @()
	$objMetaInfo +="tableHeader=$title"
	$objMetaInfo +="introduction=None found."
	$objMetaInfo +="chartable=false"
	$objMetaInfo +="titleHeaderType=h$($headerType+1)"

	return $objMetaInfo
}
function getShortDate()
{
	get-date -f "dd-MM-yyyy"
}
function getShortDateTime()
{
	get-date -f "dd-MM-yyyy HH:mm:ss"
}
function getLongDateTime()
{
	get-date -f "dddd dd-MM-yyyy H:mm:ss"
}

<#
		function declarations
#>
function setLoggingDirectory([Parameter(Mandatory=$true)][string]$dirPath)
{
	if (!$dirPath)
	{
		$global:logDir="$([IO.Path]::DirectorySeparatorChar)output"
	} else {
		$global:logDir=$dirPath
	}
}
<# Duplicates#>
<#
function showError ([Parameter(Mandatory=$true)][string] $msg, $errorColor=$global:colours.Error)
{
	#logThis ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $errorColor
	#logThis ">> " -ColourScheme $errorColor
	logThis ">> $msg" -ColourScheme $errorColor
	#logThis ">> " -ColourScheme $errorColor
	#logThis ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $errorColor
}

function verboseThis ([Parameter(Mandatory=$true)][object] $msg, $colour=$global:colours.Highlight)
{
	#logThis ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $colour
	logThis ">> " -ColourScheme $colour
	logThis ">> $msg" -ColourScheme $colour
	#logThis ">> " -ColourScheme $colour
	#logThis ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $colour
}
#>
function getRuntimeLogFileContent()
{
	if ( (Get-Variable runtimeLogFile -Scope Global) -and $global:runtimeLogFile -and (Test-Path -Path $global:runtimeLogFile) )
	{
		get-content $global:runtimeLogFile
	} elseif ( (Get-Variable logInMemory -Scope Global) -and $global:logInMemory -and (Test-Path -Path $global:logInMemory) )
	{
		Get-Content $global:logInMemory
	} else {
		return "No logs were collected during this runtime or the logs are inaccessible from this function"
	}
	#Get-Content $logFile
}


function printToScreen([string]$msg,[string]$ForegroundColor=$global:colours.Information)
{
	if (!$global:silent)
	{
		logThis -msg $msg -ColourScheme $ForegroundColor
	}
}

<#function Get-myCredentials (
			[Parameter(Mandatory=$true)][string]$User,
		  	[Parameter(Mandatory=$true)][string]$SecureFileLocation)
{
	$password = Get-Content $SecureFileLocation | ConvertTo-SecureString
	$credential = New-Object System.Management.Automation.PsCredential($user,$password)
	if ($credential)
	{
		return $credential
	} else {
		return $null
	}
}
#>

function Get-myCredentials (
	[Parameter(Mandatory=$true)][string]$User,
	[Parameter(Mandatory=$true)][string]$SecureFileLocation,
	[Parameter(Mandatory=$false)][Object]$Key=(3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43))
{
	if (Test-Path -Path $SecureFileLocation) {
		$encrypContent = Get-Content $SecureFileLocation
		if ($key)
		{
			$password =  ConvertTo-SecureString $encrypContent #-key $key
		} else {
			$password =  ConvertTo-SecureString $encrypContent
		}
		return (New-Object System.Management.Automation.PsCredential($user,$password))
	} else {
		showError -msg "Invalid secureFileLocation $SecureFileLocation"
	}
}

function getmycredentialsfromFile (
	[Parameter(Mandatory=$true)][string]$User,
	[Parameter(Mandatory=$true)][string]$SecureFileLocation,
	[Parameter(Mandatory=$false)][Object]$Key=(3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43))
{
	#logThis -msg $SecureFileLocation
	$encrypContent = Get-Content $SecureFileLocation
	if ($key)
	{
		$password =  ConvertTo-SecureString $encrypContent -key $key
	} else {
		$password =  ConvertTo-SecureString $encrypContent
	}
	return (New-Object System.Management.Automation.PsCredential($user,$password))
}

function set-mycredentials (
	[Parameter(Mandatory=$true)][string]$filename,
	[Parameter(Mandatory=$false)][Object]$Key=(3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43))
{

	#$Credential = Get-Credential -Message "Enter your credentials for this connection: "
	# old versions of powershell can't display custom messages
	$Credential = Get-Credential #-Message "Enter your credentials for this connection: "
	if ($key)
	{
		ConvertFrom-SecureString $credential.Password | Set-Content $filename
	} else {
		ConvertFrom-SecureString $credential.Password -Key $key | Set-Content $filename
	}
}

########################################################################################
# examples
# $card = Get-HashtableAsObject @{
#	Card = {2..9 + "Jack", "Queen", "King", "Ace" | Get-Random}
#	Suit = {"Clubs", "Hearts", "Diamonds", "Spades" | Get-Random}
# }
# $card
#
# $userInfo = @{
#    LocalUsers = {Get-WmiObject "Win32_UserAccount WHERE LocalAccount='True'"}
#    LoggedOnUsers = {Get-WmiObject "Win32_LoggedOnUser" }
# }
# $liveUserInfo = Get-HashtableAsObject $userInfo
function Get-HashtableAsObject([Hashtable]$hashtable)
{
    #.Synopsis
    #    Turns a Hashtable into a PowerShell object
    #.Description
    #    Creates a new object from a hashtable.
    #.Example
    #    # Creates a new object with a property foo and the value bar
    #    Get-HashtableAsObject @{"Foo"="Bar"}
    #.Example
    #    # Creates a new object with a property Random and a value
    #    # that is generated each time the property is retreived
    #    Get-HashtableAsObject @{"Random" = { Get-Random }}
    #.Example
    #    # Creates a new object from a hashtable with nested hashtables
    #    Get-HashtableAsObject @{"Foo" = @{"Random" = {Get-Random}}}
    process {
        $outputObject = New-Object Object
        if ($hashtable -and ($hashtable | measure-object).Count) {
            $hashtable.GetEnumerator() | Foreach-Object {
                if ($_.Value -is [ScriptBlock]) {
                    $outputObject = $outputObject | Add-Member ScriptProperty $_.Key $_.Value -passThru
                } else {
                    if ($_.Value -is [Hashtable]) {
                        $outputObject = $outputObject | Add-Member NoteProperty $_.Key (Get-HashtableAsObject $_.Value) -passThru
                    } else {
                        $outputObject = $outputObject | Add-Member NoteProperty $_.Key $_.Value -passThru
                    }
                }
            }
        }
        $outputObject
    }
}

function sendEmail
	(	[Parameter(Mandatory=$true)][string] $smtpServer,
		[Parameter(Mandatory=$true)][string] $from,
		[Parameter(Mandatory=$true)][string] $replyTo=$from,
		[Parameter(Mandatory=$true)][string] $toAddress,
		[Parameter(Mandatory=$true)][string] $subject,
		[Parameter(Mandatory=$true)][string] $body="",
		[Parameter(Mandatory=$false)][PSCredential] $credentials,
		[Parameter(Mandatory=$false)][string]$fromContactName="",
		[Parameter(Mandatory=$false)][object] $attachments # An array of filenames with their full path locations

	)
{
	logThis -msg "[$attachments]" -ColourScheme $global:colours.Change
	if (!$smtpServer -or !$from -or !$replyTo -or !$toAddress -or !$subject -or !$body)
	{
		logThis -msg "Cannot Send email. Missing parameters for this function. Note that All fields must be specified" -ColourScheme $global:colours.Error
		logThis -msg "smtpServer = $smtpServer"
		logThis -msg "from = $from"
		logThis -msg "replyTo = $replyTo"
		logThis -msg "toAddress = $toAddress"
		logThis -msg "subject = $subject"
		logThis -msg "body = $body"
	} else {
		if ($attachments)
		{

			<#$attachments | ForEach-Object {
				#logThis -msg $_ -ColourScheme $global:colours.Change
				$attachment = new-object System.Net.Mail.Attachment($_,"Application/Octet")
				$msg.Attachments.Add($attachment)
			}
			#>
			logThis -msg "Sending email with attachments"
			Send-MailMessage -SmtpServer $smtpServer -Credential $Credentials -From $from -Subject $subject -To $toAddress -BodyAsHtml $body -Attachments $(([string]$attachments).Trim() -replace ' ',',')
		} else {
			logThis -msg "Sending email without attachments"
			Send-MailMessage -SmtpServer $smtpServer -Credential $Credentials -From $from -Subject $subject -To $toAddress -BodyAsHtml $body -DeliveryNotificationOption OnFailure
		}
	}

}
function dumpXmlToDisk([Parameter(Mandatory=$true)][object]$xmlobj,[Parameter(Mandatory=$true)][string]$filepath,[Parameter(Mandatory=$false)][bool]$zipoutput=$true)
{
	if ($xmlobj.ObjectType)
	{
		logThis -msg "Dumping collection of all $($xmlobj.ObjectType) to disk @ $filepath"
	} else {
		logThis -msg "Dumping xmlObj to disk @ $filepath"
	}
	if ($xmlobj -and $filepath) { $xmlobj | Export-Clixml -Path $filepath }
	if ($zipoutput -and $filepath -and (Test-Path -Path $filepath))
	{
		if ((get-item -path $filepath).Length -le 2Gb)
		{
			$zippedXmlOutput = $filepath -replace '.xml','.zip'
			logThis -msg "ZIP: $zippedXmlOutput"
			$o = New-ZipFile -InputObject "$filepath" -ZipFilePath "$zippedXmlOutput"
		} else {
			logThis -msg "Unable to compress $filepath to $zippedXmlOutput because it is greater than 2Gb"
		}
	}
	# If DEBUGON is set to false then delete the XML copy but only do that if a zipped copy exists first.
	# DEBUG is FALSE and ZIPPED XML OUTPUT Exists and XML OUTPUT File exists
	if (!$global:DEBUGON -and (Test-Path "$zippedXmlOutput") -and (Test-Path "$filepath"))
	{
		Remove-Item "$filepath"
	} else {
		logThis -msg "$filepath not removed because of one of the following reasons: 1) DEBUG MODE is on, 2) Zip copy was not created, 3) The XML file was never written to disk in the first place."
	}
}

function New-ZipFile {
	#.Synopsis
	#  Create a new zip file, optionally appending to an existing zip...
	[CmdletBinding()]
	param(
		# The path of the zip to create
		[Parameter(Position=0, Mandatory=$true)]
		$ZipFilePath,

		# Items that we want to add to the ZipFile
		[Parameter(Position=1, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("PSPath","Item")]
		[string[]]$InputObject = $Pwd,

		# Append to an existing zip file, instead of overwriting it
		[Switch]$Append,

		# The compression level (defaults to Optimal):
		#   Optimal - The compression operation should be optimally compressed, even if the operation takes a longer time to complete.
		#   Fastest - The compression operation should complete as quickly as possible, even if the resulting file is not optimally compressed.
		#   NoCompression - No compression should be performed on the file.
		[System.IO.Compression.CompressionLevel]$Compression = "Optimal"
	)
	begin {
		# Make sure the folder already exists
		Add-Type -As System.IO.Compression.FileSystem
		[string]$File = Split-Path $ZipFilePath -Leaf
		[string]$Folder = $(if($Folder = Split-Path $ZipFilePath) { Resolve-Path $Folder } else { $Pwd })
		$ZipFilePath = Join-Path $Folder $File
		# If they don't want to append, make sure the zip file doesn't already exist.
		if(!$Append) {
			if(Test-Path $ZipFilePath) { Remove-Item $ZipFilePath }
		}
		$Archive = [System.IO.Compression.ZipFile]::Open( $ZipFilePath, "Update" )
	}
	process {
		foreach($path in $InputObject) {
			foreach($item in Resolve-Path $path) {
				# Push-Location so we can use Resolve-Path -Relative
				Push-Location (Split-Path $item)
				# This will 2116 the file, or all the files in the folder (recursively)
				foreach($file in Get-ChildItem $item -Recurse -File -Force | % FullName) {
					# Calculate the relative file path
					$relative = (Resolve-Path $file -Relative).TrimStart(".$([IO.Path]::DirectorySeparatorChar)")
					# Add the file to the zip
					$null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $file, $relative, $Compression)
				}
				Pop-Location
			}
		}
	}
	end {
		$Archive.Dispose()
		Get-Item $ZipFilePath
	}
}


function Expand-ZipFile {
	#.Synopsis
	#  Expand a zip file, ensuring it's contents go to a single folder ...
	[CmdletBinding()]
	param(
		# The path of the zip file that needs to be extracted
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=0, Mandatory=$true)]
		[Alias("PSPath")]
		$FilePath,

		# The path where we want the output folder to end up
		[Parameter(Position=1)]
		$OutputPath = $Pwd,

		# Make sure the resulting folder is always named the same as the archive
		[Switch]$Force
	)
	process {
		Add-Type -As System.IO.Compression.FileSystem
		$ZipFile = Get-Item $FilePath
		$Archive = [System.IO.Compression.ZipFile]::Open( $ZipFile, "Read" )

		# Figure out where we'd prefer to end up
		if(Test-Path $OutputPath) {
			# If they pass a path that exists, we want to create a new folder
			$Destination = Join-Path $OutputPath $ZipFile.BaseName
		} else {
			# Otherwise, since they passed a folder, they must want us to use it
			$Destination = $OutputPath
		}

		# The root folder of the first entry ...
		$ArchiveRoot = ($Archive.Entries[0].FullName -Split "/|\\")[0]

		Write-Verbose "Desired Destination: $Destination"
		Write-Verbose "Archive Root: $ArchiveRoot"

		# If any of the files are not in the same root folder ...
		if($Archive.Entries.FullName | Where-Object { @($_ -Split "/|\\")[0] -ne $ArchiveRoot }) {
			# extract it into a new folder:
			New-Item $Destination -Type Directory -Force
			[System.IO.Compression.ZipFileExtensions]::ExtractToDirectory( $Archive, $Destination )
		} else {
			# otherwise, extract it to the OutputPath
			[System.IO.Compression.ZipFileExtensions]::ExtractToDirectory( $Archive, $OutputPath )

			# If there was only a single file in the archive, then we'll just output that file...
			if(($Archive.Entries | Measure-Object).Count -eq 1) {
				# Except, if they asked for an OutputPath with an extension on it, we'll rename the file to that ...
				if([System.IO.Path]::GetExtension($Destination)) {
					Move-Item (Join-Path $OutputPath $Archive.Entries[0].FullName) $Destination
				} else {
					Get-Item (Join-Path $OutputPath $Archive.Entries[0].FullName)
				}
			} elseif($Force) {
				# Otherwise let's make sure that we move it to where we expect it to go, in case the zip's been renamed
				if($ArchiveRoot -ne $ZipFile.BaseName) {
					Move-Item (join-path $OutputPath $ArchiveRoot) $Destination
					Get-Item $Destination
				}
			} else {
				Get-Item (Join-Path $OutputPath $ArchiveRoot)
			}
		}

		$Archive.Dispose()
	}
}


function Set-Credentials ([Parameter(Mandatory=$true)][string]$File="securestring.txt")
{
	# This will prompt you for credentials -- be sure to be calling this function from the appropriate user
	# which will decrypt the password later on
	$Credential = Get-Credential
	$credential.Password | ConvertFrom-SecureString | Set-Content $File
}


# Add the aliases ZIP and UNZIP
#new-alias zip new-zipfile
#new-alias unzip expand-zipfile


function ChartThisTable( [Parameter(Mandatory=$true)][array]$datasource,
		[Parameter(Mandatory=$true)][string]$outputImageName,
		[Parameter(Mandatory=$true)][string]$chartType="line",
		[Parameter(Mandatory=$true)][int]$xAxisIndex=0,
		[Parameter(Mandatory=$true)][int]$yAxisIndex=1,
		[Parameter(Mandatory=$true)][int]$xAxisInterval=1,
		[Parameter(Mandatory=$true)][string]$xAxisTitle,
		[Parameter(Mandatory=$true)][int]$yAxisInterval=50,
		[Parameter(Mandatory=$true)][string]$yAxisTitle="Count",
		[Parameter(Mandatory=$true)][int]$startChartingFromColumnIndex=1, # 0 = All columns, 1 = starting from 2nd column, because you want to use Colum 0 for xAxis
		[Parameter(Mandatory=$true)][string]$title="EnterTitle",
		[Parameter(Mandatory=$true)][int]$width=800,
		[Parameter(Mandatory=$true)][int]$height=800,
		[Parameter(Mandatory=$true)][string]$BackColor="White",
		[Parameter(Mandatory=$true)][string]$fileType="png"
	  )
{
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	$colorChoices=@("#0000CC","#00CC00","#FF0000","#2F4F4F","#006400","#9900CC","#FF0099","#62B5FC","#228B22","#000080")

	$scriptpath = Split-Path -parent $outputImageName

	$headers = $datasource | Get-Member -membertype NoteProperty | Select-Object -Property Name

	logThis -msg  "+++++++++++++++++++++++++++++++++++++++++++" -ColourScheme $global:colours.Information
	logThis -msg  "Output image: $outputImageName" -ColourScheme $global:colours.Information

	logThis -msg  "Table to chart:" -ColourScheme $global:colours.Information
	logThis -msg  "" -ColourScheme $global:colours.Information
	logThis -msg  $datasource  -ColourScheme $global:colours.Information
	logThis -msg  "+++++++++++++++++++++++++++++++++++++++++++ " -ColourScheme $global:colours.Information

	# chart object
	$chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
	$chart1.Width = $width
	$chart1.Height = $height
	$chart1.BackColor = [System.Drawing.Color]::$BackColor

	# title
	[void]$chart1.Titles.Add($title)
	$chart1.Titles[0].Font = "Arial,13pt"
	$chart1.Titles[0].Alignment = "topLeft"

	# chart area
	$chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
	$chartarea.Name = "ChartArea1"
	$chartarea.AxisY.Title = $yAxisTitle #$headers[$yAxisIndex]
	$chartarea.AxisY.Interval = $yAxisInterval
	$chartarea.AxisX.Interval = $xAxisInterval
	if ($xAxisTitle) {
		$chartarea.AxisX.Title = $xAxisTitle
	} else {
		$chartarea.AxisX.Title = $headers[$xAxisIndex].Name
	}
	$chart1.ChartAreas.Add($chartarea)


	# legend
	$legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
	$legend.name = "Legend1"
	$chart1.Legends.Add($legend)

	# chart data series
	$index=0
	#$index=$startChartingFromColumnIndex
	$headers | ForEach-Object {
		$header = $_.Name
		if ($index -ge $startChartingFromColumnIndex)# -and $index -lt ($headers| Measure-Object).Count)
	    {
			logThis -msg  "Creating new series: $($header)"
			[void]$chart1.Series.Add($header)
			$chart1.Series[$header].ChartType = $chartType #Line,Column,Pie
			$chart1.Series[$header].BorderWidth  = 3
			$chart1.Series[$header].IsVisibleInLegend = $true
			$chart1.Series[$header].chartarea = "ChartArea1"
			$chart1.Series[$header].Legend = "Legend1"
			logThis -msg  "Colour choice is $($colorChoices[$index])"
			$chart1.Series[$header].color = "$($colorChoices[$index])"
		#   $datasource | ForEach-Object {$chart1.Series["VMCount"].Points.addxy( $_.date , ($_.VMCountorySize / 1000000)) }
			$datasource | ForEach-Object {
				$chart1.Series[$header].Points.addxy( $_.date , $_.$header )
			}
		}
		$index++;
	}
	# save chart
	$chart1.SaveImage($outputImageName,$fileType)

}


function mergeTables(
	[Parameter(Mandatory=$true)][string]$lookupColumn,
	[Parameter(Mandatory=$true)][Object]$refTable,
	[Parameter(Mandatory=$true)][Object]$lookupTable
)
{
	$dict=$lookupTable | Group-Object $lookupColumn -AsHashTable -AsString
	$additionalProps=diff ($refTable | gm -MemberType NoteProperty | Select-Object -ExpandProperty Name) ($lookupTable | gm -MemberType NoteProperty |
		select -ExpandProperty Name) |
		where {$_.SideIndicator -eq "=>"} | Select-Object -ExpandProperty InputObject
	$intersection=diff $refTable $lookupTable -Property $lookupColumn -IncludeEqual -ExcludeDifferent -PassThru
	foreach ($prop in $additionalProps) { $refTable | Add-Member -MemberType NoteProperty -Name $prop -Value $null -Force}
	foreach ($item in ($refTable | where {$_.SideIndicator -eq "=="})){
		$lookupKey=$(foreach($key in $lookupColumn) { $item.$key} ) -join ""
		$newVals=$dict."$lookupKey" | Select-Object *
		foreach ( $prop in $additionalProps){
			$item.$prop=$newVals.$prop
		}
	}
	$refTable | Select-Object * -ExcludeProperty SideIndicator
}
#
# This file contains a collection of parame
#

# Main function
#Add-pssnapin VMware.VimAutomation.Core
function getRuntimeDate()
{
	#return [date]$global:runtime

}
function getRuntimeDateString()
{
	return $global:runtime

}


# pass a date to this function, and it will return the 1st day of the month for this day.
# Meaning, if you pass a day of 10th January 2014, then the function should return 1 January 2014
function forThisdayGetFirstDayOfTheMonth([DateTime]$day)
{
	(get-date "1/$((Get-Date $day).Month)/$((Get-Date $day).Year)")
}

# pass a date to this function, and it will return the Last day of the month for this day.
# Meaning, if you pass a day of 10th January 2014, then the function should return 31 January 2014
function forThisdayGetLastDayOfTheMonth([DateTime]$day)
{
	(get-date "$([System.DateTime]::DaysInMonth((get-date $day).Year, (get-date $day).Month)) / $((get-date $day).Month) / $((get-date $day).Year) 23:59:59")
}

function getMonthYearColumnFormatted([DateTime]$day)
{
	(Get-Date $day -Format "MMM yyyy")
}

function daysSoFarInThisMonth([DateTime] $day)
{
	$day.Day
}

function createChart(
	[Object[]]$datasource,
	[string]$outputImage,
	[string]$chartTitle,
	[string]$xAxisTitle,
	[string]$yAxisTitle,
	[string]$imageFileType,
	[string]$chartType,
	[int]$width=800,
	[int]$height=600,
	[int]$startChartingFromColumnIndex=1,
	[int]$yAxisInterval=5,
	[int]$yAxisIndex=1,
	[int]$xAxisIndex=0,
	[int]$xAxisInterval
	)
{
	#logThis -msg "`tProcessing chart for $chartTitle"
	#logThis -msg "`t`t$sourceCSV $chartTitle $xAxisTitle $yAxisTitle $imageFileType $chartType"
	#if (!$outputImage)
	#{
		#$outputImage=$sourceCSV.Replace(".csv",".$imageFileType")
		#$imageFilename=$($sourceCSV |  Split-Path -Leaf).Replace(".csv",".$imageFileType");
	#}
	#Eventually change to table
	#$tableCSV=Import-Csv $sourceCSV
	if ($xAxisInterval -eq -1)
	{
		# I want to plot ALL the graphs
		$xAxisInterval = 1
	} else {
		$xAxisInterval = [math]::round(($datasource | Measure-Object).Count/7,0)
	}
	#$xAxisInterval = ($tableCSV | Measure-Object).Count-2

	$dunnowhat=generate-chartImageFile.ps1 -datasource $datasource `
							-title $chartTitle `
							-outputImageName $outputImage `
							-chartType  $chartType `
							-xAxisIndex $xAxisIndex `
							-xAxisTitle $xAxisTitle `
							-xAxisInterval $xAxisInterval `
							-yAxisIndex $yAxisIndex `
							-yAxisTitle $yAxisTitle `
							-yAxisInterval $yAxisInterval `
							-startChartingFromColumnIndex $startChartingFromColumnIndex `
							-width $width `
							-height $height `
							-fileType $imageFileType

	return $outputImage
}

function tableToHTMLTable([Parameter(Mandatory=$true)][Object]$table)
{
	$htmlTable += ($table | sort-object Category -Unique | Select-Object Category,Name,Description | ConvertTo-HTML -Fragment -As "Table") -replace "<table","$caption<table class=aITTablesytle" -replace "&lt;/li&gt;","</li>" -replace "&lt\;li&gt;","<li>" -replace "&lt\;/ul&gt;","</ul>" -replace "&lt\;ul&gt;","<ul>"  -replace "`r","<br>"
	$htmlTable += "`n"
	return $htmlTable
}

function formatHeaders([Parameter(Mandatory=$true)][string]$text)
{
	#return ((Get-Culture).TextInfo.ToTitleCase(($text -replace "_"," " -replace '\.',' ').ToLower()))
	$TextInfo = (Get-Culture).TextInfo
	return $TextInfo.ToTitleCase($text)
}

function formatHeaderString ([string]$string)
{
	return [Regex]::Replace($string, '\b(\w)', { param($m) $m.Value.ToUpper() });
}

function header([Parameter(Mandatory=$true)][string]$str,[Parameter(Mandatory=$true)][int]$headerType)
{
	switch ($headerType)
	{
		1 {return "`n"+ (header1 -string $str)}
		2 {return "`n"+ (header2 -string $str)}
		3 {return "`n"+ (header3 -string $str)}
		4 {return "`n"+ (header4 -string $str)}
		5 {return "`n"+ (header5 -string $str)}
		default {return "`ninvalid header type"}
	}
}

function header1([Parameter(Mandatory=$true)][string]$string)
{
	return "<h1>$(formatHeaderString $string)</h1>"
}

function header2([Parameter(Mandatory=$true)][string]$string)
{
	return "<h2>$(formatHeaderString $string)</h2>"
}

function header3([Parameter(Mandatory=$true)][string]$string)
{
	return "<h3>$(formatHeaderString $string)</h3>"
}
function header4([Parameter(Mandatory=$true)][string]$string)
{
	return "<h4>$(formatHeaderString $string)</h4>"
}
function header5([Parameter(Mandatory=$true)][string]$string)
{
	return "<h5>$(formatHeaderString $string)</h5>"
}

function paragraph([Parameter(Mandatory=$true)][string]$string)
{
	return "<p>$string</p>`n"
}

function htmlFooter()
{
	return "<p><small>$runtime | $global:srvconnection | generated from $env:computername.$env:userdnsdomain </small></p></body></html>"

}

function htmlHeader()
{
	#$content += (Get-content $global:htmlHeaderCSS)
	#$content += @"

	$content = @"
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
	<html><head><title>Report</title>
	<style type="text/css">
<!--
body {
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}

#report { width: 835px; }
.red td{
	background-color: red;
}
.yellow  td{
	background-color: yellow;
}
.green td{
	background-color: green;
}
table{
   border-collapse: collapse;
   border: 1px solid #cccccc;
   font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
   color: black;
   margin-bottom: 10px;
   margin-left: 20px;
   width: auto;
}
table td{
       font-size: 12px;
       padding-left: 0px;
       padding-right: 20px;
       text-align: left;
	   width: auto;
	   border: 1px solid #cccccc;
}
table th {
       font-size: 12px;
       font-weight: bold;
       padding-left: 0px;
       padding-right: 20px;
       text-align: left;
	   border: 1px solid #cccccc;
	   width: auto;
	   border: 1px solid #cccccc;
}

h1{
	clear: both;
	font-size: 160%;
}

h2{
	clear: both;
	font-size: 130%;
}

h3{
   clear: both;
   font-size: 120%;
   margin-left: 20px;
   margin-top: 30px;
   font-style: italic;
}

h3{
   clear: both;
   font-size: 100%;
   margin-left: 20px;
   margin-top: 30px;
   font-style: italic;
}

p{ margin-left: 20px; font-size: 12px; }

ul li {
	font-size: 12px;
}

table.list{ float: left; }

table.list td:nth-child(1){
       font-weight: bold;
       border-right: 1px grey solid;
       text-align: right;
}

table.list td:nth-child(2){ padding-left: 7px; }
table tr:nth-child(even) td:nth-child(even){ background: #CCCCCC; }
table tr:nth-child(odd) td:nth-child(odd){ background: #F2F2F2; }
table tr:nth-child(even) td:nth-child(odd){ background: #DDDDDD; }
table tr:nth-child(odd) td:nth-child(even){ background: #E5E5E5; }

div.column { width: 320px; float: left; }
div.first{ padding-right: 20px; border-right: 1px  grey solid; }
div.second{ margin-left: 30px; }
	-->
	</style>
	</head>
	<body>
"@
	return $content

}
function htmlHeaderPrev()
{
	return @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>Virtual Machine ""$guestName"" System Report</title>
<style type="text/css">
<!--
body {
	font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}

#report { width: 835px; }
.red td{
	background-color: red;
}
.yellow  td{
	background-color: yellow;
}
.green td{
	background-color: green;
}
table{
   border-collapse: collapse;
   border: 1px solid #cccccc;
   font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
   color: black;
   margin-bottom: 10px;
   margin-left: 20px;
   width: auto;
}
table td{
       font-size: 12px;
       padding-left: 0px;
       padding-right: 20px;
       text-align: left;
	   width: auto;
	   border: 1px solid #cccccc;
}
table th {
       font-size: 12px;
       font-weight: bold;
       padding-left: 0px;
       padding-right: 20px;
       text-align: left;
	   border: 1px solid #cccccc;
	   width: auto;
	   border: 1px solid #cccccc;
}

h1{
	clear: both;
	font-size: 160%;
}

h2{
	clear: both;
	font-size: 130%;
}

h3{
   clear: both;
   font-size: 120%;
   margin-left: 20px;
   margin-top: 30px;
   font-style: italic;
}

h3{
   clear: both;
   font-size: 100%;
   margin-left: 20px;
   margin-top: 30px;
   font-style: italic;
}

p{ margin-left: 20px; font-size: 12px; }

ul li {
	font-size: 12px;
}

table.list{ float: left; }

table.list td:nth-child(1){
       font-weight: bold;
       border-right: 1px grey solid;
       text-align: right;
}

table.list td:nth-child(2){ padding-left: 7px; }
table tr:nth-child(even) td:nth-child(even){ background: #CCCCCC; }
table tr:nth-child(odd) td:nth-child(odd){ background: #F2F2F2; }
table tr:nth-child(even) td:nth-child(odd){ background: #DDDDDD; }
table tr:nth-child(odd) td:nth-child(even){ background: #E5E5E5; }

div.column { width: 320px; float: left; }
div.first{ padding-right: 20px; border-right: 1px  grey solid; }
div.second{ margin-left: 30px; }
-->
</style>
</head>
<body>
"@
}


<# 
Used by older functions 
#>
function sanitiseTheReport([Parameter(Mandatory=$true)][Object]$tmpReport)
{
	$Members = $tmpReport | Select-Object `
	  @{n='MemberCount';e={ (($_ | Get-Member) | Measure-Object).Count }}, `
	  @{n='Members';e={ $_.PsObject.Properties | ForEach-Object { $_.Name } }}
	$AllMembers = ($Members | Sort-Object MemberCount -Descending)[0].Members

	$Report = $tmpReport | ForEach-Object {
	  ForEach ($Member in $AllMembers)
	  {
	    If (!($_ | Get-Member -Name $Member))
	    {
	      $_ | Add-Member -Type NoteProperty -Name $Member -Value ""
	    }
	  }
	  Write-Output $_
	}

	return $Report
}

<#
The following function will replace the above report eventually
#>
function sanitiseTheReportNew([Parameter(Mandatory=$true)][Object]$tmpReport,[Parameter(Mandatory=$false)][int]$lockFirstColums=1)
{
	# Determining the row with the largest columns, to use as the 
	$DEBUGGING=$false
	if ($DEBUGGING)
	{
		$saneIndexer=0
	}
	$fieldNamesPass1 = @()

	# 1st pass Build the whole thing
	$tmpReport | ForEach-Object {
		$indexRow=1
		$tmpReportRow = $_
		$allFieldNames = $tmpReportRow.psobject.Properties | select-object -ExpandProperty Name
		$allFieldNames | ForEach-Object  {
			$columnName = $_		
			if ($fieldNamesPass1 -contains $columnName)
			{
				if ($DEBUGGING)
				{
					Write-Host "$saneIndexer - $columnName already exist in columen array, Skipping..."
				}
			} else 
			{
				if ($DEBUGGING)
				{
						Write-Host "$saneIndexer - Adding $columnName ..." -ForegroundColor Yellow
				}
				$fieldNamesPass1 += $columnName
			}
			if ($DEBUGGING)
			{
				#pause
			}
			$saneIndexer++
		}
	}

	# 2nd pass; set the preferred first columnes as Do a first pass to lock in our first columns
	$fieldNamesPass2 = @()
	if ($lockFirstColums)
	{
		# Take the first row and add in the first 'lockFirstColums' fields at the beggining
		#$columns += $setFirstColumnAs
		$firstObject = $tmpReport | Select-Object -First 1 
		$firstObject.psobject.Properties | select-object -ExpandProperty Name | Select-Object -First $lockFirstColums | ForEach-Object {
			$fieldNamesPass2 += $_
		}

		# Add the missing list
		$fieldNamesPass1 | Sort-Object | ForEach-Object {
			$columnName = $_
			if ($fieldNamesPass2 -contains $columnName)
			{
				if ($DEBUGGING)
				{
					Write-Host "$saneIndexer - $columnName already exist in columen array, Skipping..."
				}
			} else 
			{
				if ($DEBUGGING)
				{
						Write-Host "$saneIndexer - Adding $columnName ..." -ForegroundColor Yellow
				}
				$fieldNamesPass2 += $columnName
			}
		}
	}


	$reportindex = 1
	$reportObject = @()
	$tmpReport | ForEach-Object {
		if ($DEBUGGING)
		{
			Write-Host "Processing $reportindex"
		}
		# For each row, build a new array
		$obj = $_
		$row = New-Object System.Object
		$fieldNamesPass2  | ForEach-Object { #| Sort-Object
			$columnName = $_ 
			if ($obj."$columnName")
			{
				$row | Add-Member -MemberType NoteProperty -Name $_ -Value $obj."$columnName"
			} else 
			{
				$row | Add-Member -MemberType NoteProperty -Name $_ -Value ""
			}
		}
		$reportObject += $row
		$reportindex++
	}

	if ($reportObject)
	{
		return $reportObject# | Select-Object "$setFirstColumnsAs",'*'
	}
	
}


# This returns the average of datasets
function getMeanValue([Parameter(Mandatory=$true)][int[]]$data)
{
	$data = $data | sort-object
	$dataCount = ($data | Measure-Object).Count
	if ($dataCount -gt 1)
	{
		return ($data | Measure-Object -Sum).Sum / ($data | Measure-Object).Count
	} else {
		return $data
	}
}

# This returns the median value in a datasets
# Use Median if there are outlier which would
function getMedianValue([Parameter(Mandatory=$true)][int[]]$data)
{
	$data = $data | sort-object
	$dataCount = ($data | Measure-Object).Count
	if ($dataCount -gt 1)
	{
		if ($dataCount%2) {
			#odd
			$medianvalue = $data[[math]::Floor($dataCount/2)]
		} else {
			#even
			$MedianValue = ($data[$dataCount/2],$data[$dataCount/2+1] | Measure-Object -Average).average
		}
		return $medianValue
	} else {
		return $data
	}
}

# Returns the value that is the most available
function getModeValue([Parameter(Mandatory=$true)][int[]]$data)
{
	$i=0
	$modevalue = @()
	foreach ($group in ($data | Group-Object | Sort-Object -Descending count)) {
		if (($group | Measure-Object).count -ge $i) {
			$i = ($group | Measure-Object).count
			$modevalue += $group.Name
		}
		else {
			break
		}
	}
	$modevalue
}


# Simple arithmetic addition to see if the value is a number
function isNumeric ([Parameter(Mandatory=$true)]$x) {
    try {
        0 + $x | Out-Null
        return $true
    } catch {
		showError -msg "$_"
        return $false
    }
}

# Enables all scripts to use the character or word that indicates that data is missing
function printNoData()
{
	return "-"
}
# a standard way to format numbers for reporting purposes
function formatNumbers (
	[Parameter(Mandatory=$true)]$var
)
{
	#logThis -msg  $("{0:n2}" -f $val)
	#logThis -msg $($var.gettype().Name)
	if ($(isNumeric $var))
	{
		return "{0:n2}" -f $var
	} else {
		return printNoData
	}
	#return "$([math]::Round($val,2))"
}

function showError ([Parameter(Mandatory=$true)][string] $msg, $errorColor=$global:colours.Error)
{
	#logThis -msg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $errorColor
	#logThis -msg "Error:" -ColourScheme $errorColor
	#logThis -msg " " -ColourScheme $errorColor
	logThis -msg "$msg" -ColourScheme $errorColor
	#logThis -msg " " -ColourScheme $errorColor
	#logThis -msg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $errorColor
}

function verboseThis ([Parameter(Mandatory=$true)][object] $msg, $errorColor=$global:colours.Highlight)
{
	#logThis -msg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $errorColor
	##logThis -msg "Warning " -ColourScheme $errorColor
	logThis -msg " " -ColourScheme $errorColor
	logThis -msg "$msg" -ColourScheme $errorColor
	#logThis -msg " " -ColourScheme $errorColor
	#logThis -msg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" -ColourScheme $errorColor
}

#Log To Screen and file
function logThis (
	[Parameter(Mandatory=$true)][string]$msg,
	[Parameter(Mandatory=$false)][object]$ColourScheme, #takes $global:colour.xxxx as input
	[Parameter(Mandatory=$false)][string]$logFile,
	[Parameter(Mandatory=$false)][string]$ForegroundColor = $global:colours.Information.Foreground,
	[Parameter(Mandatory=$false)][string]$BackgroundColor = $global:colours.Information.Background,
	[Parameter(Mandatory=$false)][bool]$logToScreen = $false,
	[Parameter(Mandatory=$false)][bool]$NoNewline = $false,
	[Parameter(Mandatory=$false)][bool]$keepLogInMemoryAlso=$false,
	[Parameter(Mandatory=$false)][bool]$showDate=$true
	)
{
	# overwrite the $ForegroundColor and $BackgroundColor if schema was provided
	# the schema to pass should be $global:colours.Error or $global:colours.Information etcc...$global:colours is defined at the tope of this module
	if ($showDate)
	{
		$msg = "$(getShortDateTime) $msg"
	}
	if ($ColourScheme)
	{
		$ForegroundColor = $ColourScheme.Foreground
		$BackgroundColor = $ColourScheme.Background
	}
	if ($global:logToScreen -or $logToScreen -and !$global:silent)
	{
		# Also verbose to screent
		if ($NoNewline)
		{
			Write-host $msg -BackgroundColor $BackgroundColor -Foreground $ForegroundColor -NoNewline;
		} else {
			Write-host $msg -BackgroundColor $BackgroundColor -Foreground $ForegroundColor;
		}
	}

	if ($global:runtimeLogFile -and !$global:lastLogEntry)
	{
		Set-Variable -Name lastLogEntry -Value ($global:runtimeLogFile -replace '.log','-lastest.log') -Scope Global
	}
	if ($global:logTofile)
	{
		if ($global:logDir -and ((Test-Path -path $global:logDir) -ne $true))
		{
			New-Item -type directory -Path $global:logDir
			$childitem = Get-Item -Path $global:logDir
			$global:logDir = $childitem.FullName
		}
		if ($logFile)
		{
			if (Test-Path -Path $logFile)
			{
				"$msg`n"  | out-file -filepath $logFile -append
			} else {
				logThis -msg "Error while writing to $logFile`n"
			}
		}
		if ($global:runtimeLogFile -and (Test-Path -Path $global:runtimeLogFile))
		{
			"$msg`n" | out-file -filepath $global:runtimeLogFile -append
		}

		if ($global:lastLogEntry -and (Test-Path -Path $global:lastLogEntry))
		{
			"$msg`n" | out-file -filepath $global:lastLogEntry
		}
	}
	if ($global:logInMemory -or $keepLogInMemoryAlso)
	{
		$global:runtimeLogFileInMemory += "$msg`n"
	}
}

<#function getRuntimeLogFileContent()
{
	if ($global:logTofile -and $global:runtimeLogFile -and (Test-Path -Path $global:runtimeLogFile)) { return get-content $global:runtimeLogFile }
	if ($global:logInMemory ) { return $global:runtimeLogFileInMemory }
	#Get-Content $logFile
}#>

function SetmyLogFile(
		[Parameter(Mandatory=$true)][string] $filename
	)
{
	if(!(Get-Variable -Name runtimeLogFile -Scope Global -ErrorAction Ignore))
	{
		Set-Variable -Name runtimeLogFile -Value $filename -Scope Global
		logThis -msg "the global:runtimeLogFile does not exist, setting it to $($global:runtimeLogFile)"
	} else {
		logThis -msg "The runtime log file is already set. Re-using and logging to $($global:runtimeLogFile)"
	}
	if (!(Test-Path -path $global:runtimeLogFile))
	{
		getLongDateTime | out-file $filename
	}
}

function SetmyCSVOutputFile(
		[Parameter(Mandatory=$true)][string] $filename
	)
{
	#logThis -msg "[SetmyCSVOutputFile] This script will log all data output to CSV file called $global:runtimeCSVOutput"
	if (!$global:runtimeCSVOutput)
	{
		Set-Variable -Name runtimeCSVOutput -Value $filename -Scope Global
	}
	if (!(Test-Path -path $global:runtimeCSVOutput))
	{
		getLongDateTime | out-file $filename
	}
}


function AppendToCSVFile (
	[Parameter(Mandatory=$true)][string] $msg
	)
{
	if ((Test-Path -path $global:logDir) -ne $true) {
		New-Item -type directory -Path $global:logDir
		$childitem = Get-Item -Path $global:logDir
		$global:logDir = $childitem.FullName
	}
	Write-Output $msg >> $global:runtimeCSVOutput
}

function ExportCSV (
	[Parameter(Mandatory=$true)][Object] $table,
	[Parameter(Mandatory=$false)][string] $sortBy,
	[Parameter(Mandatory=$false)][string] $thisFileInstead,
	[Parameter(Mandatory=$false)][object[]] $metaData
	)
{

	$report = sanitiseTheReport $table
	if ((Test-Path -path $global:logDir) -ne $true) {
		New-Item -type directory -Path $global:logDir
		$childitem = Get-Item -Path $global:logDir
		$global:logDir = $childitem.FullName
	}
	$filename=$global:runtimeCSVOutput
	if ($thisFileInstead)
	{
		$filename = $thisFileInstead
	}
	LogThis "outputCSV = $filename"
	if ($sortBy)
	{
			$report | Sort-Object -Property $sortBy -Descending | Export-Csv -Path $filename -NoTypeInformation
	} else {
			$report | Export-Csv -Path $filename -NoTypeInformation
	}

	if ($metadata)
	{
		ExportMetaData -metadata $metadata
	}
}

function returnResults (
	[Parameter(Mandatory=$true)][Object] $table,
	[Parameter(Mandatory=$true)][Object] $metaInfoTable,
	[Parameter(Mandatory=$true)][Object] $runtimeLogs,
	[Parameter(Mandatory=$false)][string] $sortBy,
	[Parameter(Mandatory=$false)][string] $thisFileInstead,
	[Parameter(Mandatory=$false)][object[]] $metaData
	)
{

	$report = sanitiseTheReport $table
	return $report,$metaInfoTable,$runtimeLogs
}


function launchReport()
{
	Invoke-Expression $global:runtimeCSVOutput
}


#######################################################################
# Logger Module - Used to log script runtime and output to screen
#######################################################################
# Source = VM (VirtualMachineImpl) or ESX objects(VMHostImpl)
# metrics = @(), an array of valid metrics for the object passed
# filters = @(), and array which contains inclusive filtering strings about specific Hardware Instances to filter the results so that they are included
# returns a script of HTML or CSV or what
function compareString([string] $first, [string] $second, [switch] $ignoreCase)
{

	# No NULL check needed
	# PowerShell parameter handling converts Nulls into empty strings
	# so we will never get a NULL string but we may get empty strings(length = 0)
	#########################

	$len1 = $first.length
	$len2 = $second.length

	# If either string has length of zero, the # of edits/distance between them
	# is simply the length of the other string
	#######################################
	if($len1 -eq 0)
	{ return $len2 }

	if($len2 -eq 0)
	{ return $len1 }

	# make everything lowercase if ignoreCase flag is set
	if($ignoreCase -eq $true)
	{
	  $first = $first.tolowerinvariant()
	  $second = $second.tolowerinvariant()
	}

	# create 2d Array to store the "distances"
	$dist = new-object -type 'int[,]' -arg ($len1+1),($len2+1)

	# initialize the first row and first column which represent the 2
	# strings we're comparing
	for($i = 0; $i -le $len1; $i++)
	{  $dist[$i,0] = $i }
	for($j = 0; $j -le $len2; $j++)
	{  $dist[0,$j] = $j }

	$cost = 0

	for($i = 1; $i -le $len1;$i++)
	{
	  for($j = 1; $j -le $len2;$j++)
	  {
	    if($second[$j-1] -ceq $first[$i-1])
	    {
	      $cost = 0
	    }
	    else
	    {
	      $cost = 1
	    }

	    # The value going into the cell is the min of 3 possibilities:
	    # 1. The cell immediately above plus 1
	    # 2. The cell immediately to the left plus 1
	    # 3. The cell diagonally above and to the left plus the 'cost'
	    ##############
	    # I had to add lots of parentheses to "help" the Powershell parser
	    # And I separated out the tempmin variable for readability
	    $tempmin = [System.Math]::Min(([int]$dist[($i-1),$j]+1) , ([int]$dist[$i,($j-1)]+1))
	    $dist[$i,$j] = [System.Math]::Min($tempmin, ([int]$dist[($i-1),($j-1)] + $cost))
	  }
	}

	# the actual distance is stored in the bottom right cell
	return $dist[$len1, $len2];
}

function setSectionHeader (
		[Parameter(Mandatory=$true)][ValidateSet('h1','h2','h3','h4','h5')][string]$type="h1",
		[Parameter(Mandatory=$true)][object]$title,
		[Parameter(Mandatory=$false)][object]$text
	)
{
	$csvFilename=(getRuntimeCSVOutput) -replace ".csv","-$($title -replace ' ','_').csv"
	$metaFilename=$csvFilename -replace '.csv','.nfo'
	$metaInfo = @()
	$metaInfo +="tableHeader=$title $SHOWCOMMANDS"
	$metaInfo +="titleHeaderType=$type"
	if ($text) { $metaInfo +="introduction=$text" }
	#$metaInfo +="displayTableOrientation=$displayTableOrientation"
	#$metaInfo +="chartable=false"
	#$metaInfo +="showTableCaption=$showTableCaption"
	#if ($metaAnalytics) {$metaInfo += $metaAnalytics}
	#ExportCSV -table $dataTable -thisFileInstead $csvFilename
	ExportMetaData -metadata $metaInfo -thisFileInstead $metaFilename
	updateReportIndexer -string "$(split-path -path $csvFilename -leaf)"
}

function getDate()
{
	if ($global.dateUFormat)
	{
		logThis -msg "$($global.dateUFormat)"
		(get-date -UFormat $global.dateUFormat)
	} else {
		(get-date)
	}
}

function convertEpochDate([Parameter(Mandatory=$true)][double]$sec)
{
	$dayone=get-date "1-Jan-1970"
	(get-date $dayone.AddSeconds([double]$sec) -format "dd-MM-yyyy hh:mm:ss")
}

function getTimeSpanFormatted([Parameter(Mandatory=$true)][TimeSpan]$timespan)
{
	$timeTakenStr=""
	if ($timespan.Days -gt 0)
	{
		$timeTakenStr += "$($timespan.days) days "
	}
	if ($timespan.Hours -gt 0)
	{
		$timeTakenStr += "$($timespan.Hours) hrs "
	}
	if ($timespan.Minutes -gt 0)
	{
		$timeTakenStr += "$($timespan.Minutes) min "
	}
	if ($timespan.Seconds -gt 0)
	{
		$timeTakenStr += "$($timespan.Seconds) sec "
	}
	return $timeTakenStr
}

function Convert-UTCtoLocal([parameter(Mandatory=$true)][String]$UTCTime)
{
	try {
		$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName;
		$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone);
		$LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ);
		return $LocalTime;
	}
	catch {
		return $null;
	}
}

# this function intakes an array of strings and returns a comma separated list
#
function returnListAsString ([Parameter(Mandatory=$true)][Object]$list,[Parameter(Mandatory=$false)][string]$separator=',')
{
	$index=0
	$returnedList=""
	$count = ($list | Measure-Object).Count
	if ($count -eq 1)
	{
		$list
	} else {
		while($index -lt $count) {
			if ($list[$index])
			{
				$returnedList += $list[$index] + $separator
			}
			$index++
		}
		# remove the last $separator
		$returnedList = $returnedList -replace ",$",""
		$returnedList
	}
}

# Syntax getSize -unit <CurrentUnit_fortheVAlue_example_KB> -val <value>
# getSize -unit "KB" -val 1024 -> returns 1MB
# getSize -unit "GB" -val 1024 -> returns 1TB
function getSize($TotalKB,$unit,$val)
{

	$valInt = [int]($val -replace ',','')
	if ($TotalKB) { $unit="KB"; $val=$TotalKB}

	if ($unit -eq "B") { $bytes=$valInt}
	elseif ($unit -eq "KB") { $bytes=$valInt*1KB }
	elseif ($unit -eq "MB") { $bytes=$valInt*1MB }
	elseif ($unit -eq "GB") { $bytes=$valInt*1GB }
	elseif ($unit -eq "TB") { $bytes=$valInt*1TB }
	elseif ($unit -eq "GB") { $bytes=$valInt*1PB }

	If ($bytes -lt 1MB) # Format TotalKB to reflect:
    {
     $value = "{0:N} KB" -f $($bytes/1KB) # KiloBytes or,
    }
    If (($bytes -ge 1MB) -AND ($bytes -lt 1GB))
    {
     $value = "{0:N} MB" -f $($bytes/1MB) # MegaBytes or,
    }
    If (($bytes -ge 1GB) -AND ($bytes -lt 1TB))
     {
     $value = "{0:N} GB" -f $($bytes/1GB) # GigaBytes or,
    }
    If ($bytes -ge 1TB -and $bytes -lt 1PB)
    {
     $value = "{0:N} TB" -f $($bytes/1TB) # TeraBytes
    }
	If ($bytes -ge 1PB)
  	 {
		$value = "{0:N} PB" -f $($bytes/1PB) # TeraBytes
    }
	return $value
}

function getSpeed($unit="KBps", $val) #$TotalBps,$kbps,$TotalMBps,$TotalGBps,$TotalTBps,$TotalPBps)
{
	$valInt = [int]($val -replace ',','')
	if ($unit -eq "bps") { $bytesps=$valInt }
	elseif ($unit -eq "kbps") { $bytesps=$valInt*1KB }
	elseif ($unit -eq "mbps") {  $bytesps=$valInt*1MB }
	elseif ($unit -eq "gbps") { $bytesps=$valInt*1GB }
	elseif ($unit -eq "tbps") { $bytesps=$valInt*1TB }
	elseif ($unit -eq "pbps") { $bytesps=$valInt*1PB }

	If ($bytesps -lt 1MB) # Format TotalKB to reflect:
    {
     $value = "{0:N} KBps" -f $($bytesps/1KB) # KiloBytes or,
    }
    If (($bytesps -ge 1MB) -AND ($bytesps -lt 1GB))
    {
     $value = "{0:N} MBps" -f $($bytesps/1MB) # MegaBytes or,
    }
    If (($bytesps -ge 1GB) -AND ($bytesps -lt 1TB))
     {
     $value = "{0:N} GBps" -f $($bytesps/1GB) # GigaBytes or,
    }
    If ($bytesps -ge 1TB -and $bytesps -lt 1PB)
    {
     $value = "{0:N} TBps" -f $($bytesps/1TB) # TeraBytes
    }
	 If ($bytesps -ge 1PB)
    {
     $value = "{0:N} PBps" -f $($bytesps/1PB) # TeraBytes
    }
	return $value
}
#getSpeed -unit $unit 100
function convertValue($unit,$val)
{
	switch($unit)
	{
		"%"    { $value = "{0:N} %" -f $val; $type="perc" }
		"bps"  { $bytes=$val; type="speed" }
		"kbps" { $bytes=$val*1KB; type="speed" }
		"mbps" {  $bytes=$val*1MB; type="speed" }
		"gbps" { $bytes=$val*1GB; type="speed" }
		"tbps" { $bytes=$val*1TB; type="speed" }
		"pbps" { $bytes=$val*1PB; type="speed" }
		"bytes"{ $bytes=$val; type="size" }
		"KB" { $bytes=$val*1KB; type="size"}
		"MB" { $bytes=$val*1MB; type="size"}
		"GB" { $bytes=$val*1GB; type="size"}
		"TB" { $bytes=$val*1TB; type="size"}
		"PB" { $bytes=$val*1PB; type="size"}
		"Hz" { $bytes=$val; type="frequency"}
		"Khz" { $bytes=$val*1000; type="frequency"}
		"Mhz" { $bytes=$val*1000*1000; type="frequency"}
		"Ghz" { $bytes=$val*1000*1000*1000; type="frequency"}
		"Thz" { $bytes=$val*1000*1000*1000*1000; type="frequency"}
	}
}

function SetmyCSVMetaFile(
	[Parameter(Mandatory=$true)][string] $filename
	)
{
	if(!$global:runtimeCSVMetaFile)
	{
		Set-Variable -Name runtimeCSVMetaFile -Value $filename -Scope Global
	}
	if (!(Test-Path -path $global:runtimeCSVMetaFile))
	{
		get-date | out-file $filename
	}
}

###############################################
# CSV META FILES
###############################################
# Meta data needed by the porting engine to
# These are the available fields for each report
#$metaInfo = @()
function New-MetaInfo {
    param(
        [Parameter(Mandatory=$true)][string]$file,
		[Parameter(Mandatory=$true)][string]$TableHeader,
		[Parameter(Mandatory=$false)][string]$Introduction,
        [Parameter(Mandatory=$false)][ValidateSet('h1','h2','h3','h4','h5',$null)][string]$titleHeaderType="h1",
		[Parameter(Mandatory=$false)][ValidateSet('false','true')][string]$TableShowCaption='false',
		[Parameter(Mandatory=$false)][ValidateSet('Table','List')][string]$TableOrientation='Table',
		[Parameter(Mandatory=$false)]$displaytable="true",
		[Parameter(Mandatory=$false)]$TopConsumer=10,
		[Parameter(Mandatory=$false)]$Top_Column,
		[Parameter(Mandatory=$false)]$chartStandardWidth="800",
		[Parameter(Mandatory=$false)]$chartStandardHeight="400",
		[Parameter(Mandatory=$false)][ValidateSet('png')]$chartImageFileType="png",
		[Parameter(Mandatory=$false)][ValidateSet('StackedBar100')]$chartType="StackedBar100",
		[Parameter(Mandatory=$false)]$chartText,
		[Parameter(Mandatory=$false)]$chartTitle,
		[Parameter(Mandatory=$false)]$yAxisTitle="%",
		[Parameter(Mandatory=$false)]$xAxisTitle="/",
		[Parameter(Mandatory=$false)]$startChartingFromColumnIndex=1,
		[Parameter(Mandatory=$false)]$yAxisInterval=10,
		[Parameter(Mandatory=$false)]$yAxisIndex=1,
		[Parameter(Mandatory=$false)]$xAxisIndex=0,
		[Parameter(Mandatory=$false)]$xAxisInterval=-1,
		[Parameter(Mandatory=$false)][Object]$Table
	)
    New-Object psobject -property @{
        file = $file
		Table = $null
		TableHeader = $TableHeader
		TableOrientation = $displayTableOrientation
		TableShowCaption = $TableShowCaption
		Introduction = $Introduction
		titleHeaderType = $titleHeaderType
		displaytable = $displaytable
		generateTopConsumers = $generateTopConsumers
		generateTopConsumersSortByColumn = $generateTopConsumersSortByColumn
		chartStandardWidth = $chartStandardWidth
		chartStandardHeight = $chartStandardHeight
		chartImageFileType = $chartImageFileType
		chartType = $chartType
		chartText = $chartText
		chartTitle = $chartTitle
		yAxisTitle = $yAxisTitle
		xAxisTitle = $xAxisTitle
		startChartingFromColumnIndex = $startChartingFromColumnIndex
		yAxisInterval = $yAxisInterval
		yAxisIndex = $yAxisIndex
		xAxisIndex= $xAxisIndex
		xAxisInterval = $xAxisInterval

    }
}

function ExportMetaData([Parameter(Mandatory=$true)][object[]] $metaData, [Parameter(Mandatory=$false)]$thisFileInstead)
{
	if ((Test-Path -path $global:logDir) -ne $true) {
		New-Item -type directory -Path $global:logDir
		$childitem = Get-Item -Path $global:logDir
		$global:logDir = $childitem.FullName
	}

	$tmpMetafile = $global:runtimeCSVMetaFile
	if( $thisFileInstead)
	{
		$tmpMetafile = $thisFileInstead
	}
	#if ($global:runtimeCSVMetaFile)
	if ($global:logTofile -and $tmpMetafile)
	{
		 $metadata | Out-File -FilePath $tmpMetafile
	}
}
function getRuntimeMetaFile()
{
	return "$global:runtimeCSVMetaFile"
}
function generateRuntimeCSVOuput([Parameter(Mandatory=$true)][string]$vcenterName,[Parameter(Mandatory=$true)][string]$scriptname,[Parameter(Mandatory=$true)][string]$logDir,[Parameter(Mandatory=$false)][string]$comment)
{
	$runtime="$(get-date -f dd-MM-yyyy)"	
	if ($comment -eq "" ) {
		$of = "$($logDir)$($([IO.Path]::DirectorySeparatorChar))$($runtime)_$($vcenterName.ToUpper())_$scriptName.csv"
	} else {
		$of = "$($logDir)$($([IO.Path]::DirectorySeparatorChar))$($runtime)_$($vcenterName.ToUpper())_$scriptName_$comment.csv"
	}
	return $of

}

function setRuntimeMetaFile([string]$filename)
{
	$global:runtimeCSVMetaFile = $filename
}

function updateRuntimeMetaFile([object[]] $metaData)
{
	if ($global:logTofile) { $metadata | Out-File -FilePath $global:runtimeCSVMetaFile -Append }
}

function getRuntimeCSVOutput()
{
	return "$global:runtimeCSVOutput"
}

function setReportIndexer($fullName)
{
	Set-Variable -Name reportIndex -Value $fullName -Scope Global
}

function getReportIndexer()
{
	if ($global:reportIndex) { return $global:reportIndex}
}

function updateReportIndexer($string)
{
	if ($global:logTofile)
	{
		$string -replace ".csv",".nfo" -replace ".ps1",".nfo" -replace ".log",".nfo" | Out-file -Append -FilePath $global:reportIndex
	}
}

#######################################################################################################
# Generate Reports from XML to EXCEL - To be completed
#######################################################################################################
function generateExcelReport(
	[Parameter(Mandatory=$true)]$xml,
	[Parameter(Mandatory=$false)]$outputXLSXfile)
{
	logThis -msg "Generating CSVs first"
	$fileList = generateCsvReport -xml $xml
	if (!$outputXLSXfile)
	{
		$outputXLSXFile = "$($global:logDir)\Reports-"+ ($($xml.Runtime.StartTime) -replace '/'.'-' -replace ':','-' -replace ' ','_') + ".xlsx"
	}
	$files = $fileList | ForEach-Object {
		$fileItem = $_ | Get-Item  | sort-object -property Name
		logThis -msg "File list = $($_) is Filename=$($fileItem.Name)"
		$fileItem
	}

	logThis -msg "Generating Excel from CSVs"
	$fileCount = ($files | Measure-Object).Count;
	logThis -msg "Number of CSV files found in folder $inputDirectory (count = $fileCount)" -ColourScheme $global:colours.Highlight;
	if ($fileCount -le 0)
	{
		logThis -msg "No input files (*.CSV) found in folder $inputDirectory (count = $fileCount)" -ColourScheme $global:colours.Highlight;
	}
	else
	{
		$excelApp = New-Object -com Excel.Application
		#$excelApp.Visible = $True
		$excelApp.DisplayAlerts = $false
		$excelApp.Visible = $true
		$book = $excelApp.Workbooks.Add()

		$fileNum = 1
		$files | ForEach-Object {
			$file = $_
			Write-Progress -Id 1 -Activity "Processing CSV  ""$($file.name)""" -CurrentOperation "$fileNum/$($fileCount)" -PercentComplete $($fileNum/$($fileCount)*100)
			#$sheetName = formatHeaders -text $($file.Name -replace ".csv",'' -replace "command","Cmd" -replace "Hardware","HW" -replace "OperatingSystem","OS" -replace "Version","Ver" -replace "Application","App" -replace "Cluster","Clus" -replace "Capacity","Cap" -replace "SoftwareDefinedDataCenter","SDDC" -replace "average","Avg" -replace "maximum","Max" -replace "minimum","Min" -replace "summation","Sum" -replace "provisioned","Prov" -replace "ratio","%" -replace "datastore","Dstore" -replace "Reservation","Reverv" -replace "Number","Num" -replace "virtualDisk","VDisk" -replace "count","Count" -replace "VirtualMachine","VM" -replace "Infrastructure","Inf" -replace "Overview","" -replace '-','' -replace "Performance","Perf" -replace "Memory","MEM" -replace "License","Lic" -replace '_','' -replace "Hypervisor","ESX" -replace "ResourceUsage","" -replace "Services","Svc" -replace "Effective","Efect")
			# Need to massage the file name into an appropriate sheet label so excel will create the worksheet
			$sheetName = $($file.Name -replace ".csv",'' -replace "command","Cmd" -replace "Hardware","HW" -replace "OperatingSystem","OS" -replace "Version","Ver" -replace "Application","App" -replace "Cluster","Clu" -replace "Capacity","Cap" -replace "SoftwareDefinedDataCenter","SDDC" -replace "average","Avg" -replace "maximum","Max" -replace "minimum","Min" -replace "summation","Sum" -replace "provisioned","Prov" -replace "ratio","%" -replace "datastore","Dstore" -replace "Reservation","Reverv" -replace "Number","Num" -replace "virtualDisk","VDisk" -replace "count","Count" -replace "VirtualMachine","VM" -replace "Infrastructure","Inf" -replace "Overview","" -replace '-','' -replace "Performance","Perf" -replace "Memory","MEM" -replace "License","Lic" -replace '_','' -replace "Hypervisor","ESX" -replace "ResourceUsage","" -replace "Services","Svc" -replace "Effective","Efect" -replace "Network","Net" -replace "net","Net" -replace "mem","Mem" -replace "cpu","Cpu" -replace "disk","Disk" -replace "usage","Usage" -replace "consumed","Consumed" -replace "throughput","Thru" -replace "mhz","Mhz" -replace "active","Active"  -replace "demand","Demand" -replace "total","Total" -replace "totalmb","TotalMB" -replace "swap","Swap" -replace "Manufacturer","Manu" -replace "used","Used" -replace "contention","Contention" -replace "by","By")
			logThis -msg "Filename $($file.Name) $sheetName"

			if ($sheetName.Length -gt 31)
			{
				logThis -msg "Sheetname longer than 31 characters: Before - $sheetName"
				$sheetName = $sheetName.Substring(0,31)
				logThis -msg "Sheetname longer than 31 characters: AFTER - $sheetName"
			}
			$sheetIndex = $fileNum - 1

			if ( $sheetIndex -ge ($book.Worksheets.Item | Measure-Object).Count) {
				logThis -msg "Adding sheet $filenum $($file.Name)"
				$currSheet = $book.Worksheets.Add()
			} else{
				logThis -msg "Reusing sheet ($filenum) $($file.name)"
				$currSheet = $book.Worksheets.Item($filenum)
			}

			if ($currSheet)
			{
				logThis -msg "`tRenaming sheet from $($currSheet.Name) to $sheetName"
				$currSheet.Activate() | Out-Null # not sure what this does
				try
				{
					logthis -msg ">>Sheetname: $sheetName<<"
					$currSheet.Name = $sheetName
				}catch{
					showError -msg ">>Sheetname: $sheetName<<"
					showError -msg "$_"
#					Write-error $_
				}


				# Open up the CSV and copy the content of the file its content
				# to our main active worksheet using paste
				$tempBook = $excelApp.Workbooks.Open($file);
				$tempsheet = $tempBook.Worksheets.Item(1);
				#Copy contents of the CSV file
				$tempSheet.UsedRange.Copy() | Out-Null
				#Paste contents of CSV into existing workbook
				$currSheet.Paste()  | Out-Null
				$tempBook.Close($false)  | Out-Null # Closing the temporary Excel window opened to read in the CSV

				# Now position the curesor into the main EXCEL spreadsheet to paste in our content
				$table = $currSheet.UsedRange;
				$table.Name = $currSheet.Name;
				$table.AutoFormat() | Out-Null;
				$table.Style.ShrinkToFit = $true;
				$table.Columns.AutoFilter() | Out-Null;
				$table.Columns.AutoFit() | Out-Null;
				$table.Style.Interior.Color = 16777215;
				$table.Style.Interior.Pattern = 0;
				$table.Style.Interior.ThemeColor = -4142
			}

			$fileNum++;
		}

		$book.SaveAs($outputXLSXFile)
		logThis -msg "Reports saved to $outputXLSXFile"
		$book.Close($false)
		$excelApp.Quit()
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
		Remove-Variable excelApp
		# delete the CSVs
		$fileList | Get-Item | Remove-Item
	}
}


#######################################################################################################
# Generate Reports from XML to EXCEL - To be completed
#######################################################################################################
# Report structure
#xml.<category>.Reports.<type>.<report>
# Example:
#              $xml.VMware(HashTable).Reports.Capacity(HashTable).Report1(HashTable).MetaData(HashTable)
#              $xml.VMware(HashTable).Reports.Capacity(HashTable).Report1(HashTable).DataTable(Object[])
#              $xml.VMware(HashTable).Reports.Capacity(HashTable).Report1(HashTable).Runtime(Object[])
#              $xml.VMware(HashTable).Reports.Issues(HashTable).(HashTable).MetaData(HashTable)
#              $xml.VMware(HashTable).Reports.Issues(HashTable).Report1(HashTable).DataTable(Object[])
#              $xml.VMware(HashTable).Reports.Issues(HashTable).Report1(HashTable).Runtime(Object[])
#
function generateCsvReport(
			[Parameter(Mandatory=$true)]$xml)
{
	logThis -msg "Generating CSV reports"

	if ($xml.keys)
	{
		$csvfilesExported=@()
		$reportTitles = $xml.keys | Where-Object {$_ -ne "Runtime"}
		logThis -msg "Keys = $([string]$reportTitles)"
		$index=1
		$reportTitles | ForEach-Object {
			#$subReportXML = $_
			$reportTitle = $_
			Write-Progress -Id 1 -Activity "Processing $reportTitle" -CurrentOperation "$index/$(($reportTitles | Measure-Object).Count)" -PercentComplete $($index/$(($reportTitles | Measure-Object).Count)*100)
			logThis -msg "[$index / $(($reportTitles | Measure-Object).Count)] reportTitle = $_"

			#LogThis -msg "$($xml.$reportTitle.
			$subReportsTitles = $xml.$reportTitle.Reports.keys
			#logThis -msg $subReportsTitles
			$subReportsTitles | ForEach-Object {
				$subReportsTitle = $_
				$csvfilename +=  ($("$reportTitle-$subReportsTitle") -replace ' ',')')
				logThis -msg "subReportsTitle = $subReportsTitle"
				for ($jindex = 0 ; $jindex -lt ($xml.$reportTitle.Reports.$subReportsTitle | Measure-Object).Count; $jindex++)
				{
					# Check if the subreport has a datatable. if NOT there is no point printing the section's other information.
					if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].DataTable)
					{
						if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].MetaData)
						{
							$metaData = convertTextVariablesIntoObject ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].MetaData)
						}
						$reportTables = $xml.$reportTitle.Reports.$subReportsTitle[$jindex].DataTable
						logThis -msg ">> [$($jindex+1)/$(($xml.$reportTitle.Reports.$subReportsTitle | Measure-Object).Count)] $($metaData.tableHeader)"

						$expectedPassesToMakeForTopConsumers = 2
						if ($metaData.generateTopConsumers)
						{
							# Report specifies that the report should generate a Top 10 (or other NUMBER) for this report, so need to make a second pass for the Top consumer report
							$passesToMake = $expectedPassesToMakeForTopConsumers
						} else {
							$passesToMake = 1
						}

						$passesSoFar=1;
						# If the report NFO file
						while ($passesSoFar -le $passesToMake)
						{
							if ($passesSoFar -eq $expectedPassesToMakeForTopConsumers)
							{
								$extraText = "(Top ($metaData.generateTopConsumers))"
							} else {
								$extraText = ""
							}

							if ($metaData.tableHeader)
							{
								# Rewrite the csvfilename to match the table header
								$csvfilenameTitle = ("$($subReportsTitle)-$($metaData.tableHeader)" -replace ' ','')
								logThis -msg "$csvfilenameTitle"
								#logThis -msg "`t-> Using heading from NFO: $($metaData.tableHeader) for CSV File name"  #-ColourScheme $color
							} else {
								#logThis -msg "`t-> Will derive Heading from the report" #-ColourScheme $color
								$csvfilenameTitle = "$title$extraText"

							}
							# sometime the dataTable in those reports can be a straigh forward Table (custom object) but at times, particularly
							# for the Performance results, it can be a hashtable with tables below for each metric. For this reason
							# i need to create a hashtable so I can feed it into the next phase.
							if ($reportTables -and ($reportTables.gettype().Name -eq "HashTable"))
							{
								#logThis -msg "This is HashTable"
								# The fact that this DataTable is a Hashtable tells me that this table is actually a table of datatable (performance reports are done this way)
								$tablesForProcessing = $reportTables
							} else {
								$csvfilenameTitle = ("$($subReportsTitle)" -replace ' ','')
								# The fact that this DataTable is NOT a Hashtable tells me that this table is a custom object I created which is a results of a table
								logThis -msg "In here: $($metaData.tableHeader)"
								if ($metaData.tableHeader)
								{
									$tablesForProcessingName = $metaData.tableHeader -replace ' ',''
									$tablesForProcessing = @{$tablesForProcessingName=$reportTables}
								} else {
									$tablesForProcessing = @{"dataTable$jindex"=$reportTables}
								}

							}
							$dataTableKeys = $tablesForProcessing.Keys
							$dataTableKeys | ForEach-Object {
								#logThis -msg "Value $($_.value.name)"
								$dataTableKeyName=$_
								$dataTable = $tablesForProcessing[$dataTableKeyName]
								if($dataTable)
								{
									if ($passesSoFar -eq $expectedPassesToMakeForTopConsumers)
									{
										if ($metaData.generateTopConsumersSortByColumn)
										{
											$tmpreport = $dataTable | Sort-Object -Property $metaData.generateTopConsumersSortByColumn -Descending | Select-Object -First $metaData.generateTopConsumers
											$dataTable = $tmpreport
										} else {
											$tmpreport = $dataTable | sort-object | Select-Object -First $metaData.generateTopConsumers
											$dataTable = $tmpreport
										}
									}
									# Export table to CSV
									#$csvFilename = "$($global:logDir)\$($csvfilenameTitle)-$($dataTableKeyName -replace '.','_').csv"
									$dataTableKeyNameTitle = $dataTableKeyName -replace '\.','_'
									$csvFilename = "$($global:logDir)\$($csvfilenameTitle)-$($dataTableKeyNameTitle).csv"
									logThis -msg "Exporting DataTable to $csvFilename"
									$dataTable | Export-Csv -NoTypeInformation -Path "$csvFilename"
									$csvfilesExported += "$csvFilename"
									#Remove-Variable $report
								} else {
									logThis -msg "No DataTable found in this subreport" -ColourScheme $global:colours.Information
								}
							}
							$passesSoFar++
						}
					}

				}
			}
			$index++
		}
		return $csvfilesExported
	} else {
		logThis -msg "Invalid input xml file"
	}
}

#######################################################################################################
# This function takes in a Array table, and transposes the Colum heads to become rows and row headers to become the Column.
#######################################################################################################
function transposeTable ([Parameter(Mandatory=$true)][object]$tableIn)
{
	$tableOut = @()
    foreach ($Property in $tableIn.Property | Select-Object -Unique) {
		$Props = [ordered]@{ Property = $Property }
		foreach ($Server in $tableIn.Server | Select-Object -Unique)
		{
			$Value = ($tableIn.where({ $_.Server -eq $Server -and $_.Property -eq $Property })).Value
			$Props += @{ $Server = $Value }
		}
		$tableOut += New-Object -TypeName PSObject -Property $Props
		return $tableOut
	}
}

#######################################################################################################
# Generate Reports from XML to HTML. The XML must be structured this way
# examples:
# 		$xml.VMware.Reports.Capacity.0.DataTable < contains the dataTable or table to write in HTML <table></table>
#			$xml.VMware.Reports.Capacity.0.MetaData <- must contain "tableHeader" and should contain Introduction="Describes the section" and titleHeaderType=h1
# 		$xml.San.Reports.Capacity[0].DataTable
#			$xml.San.Reports.Capacity[0].MetaData
# 		$xml.Backups.Reports.Capacity[0].DataTable
#			$xml.Backups.Reports.Capacity[0].MetaData
#
#
#
#######################################################################################################
function generateHTMLReport(
			[Parameter(Mandatory=$true)][string]$reportHeader,
			[Parameter(Mandatory=$true)][string]$reportIntro,
			[Parameter(Mandatory=$true)][string]$farmName,
			[Parameter(Mandatory=$False)][bool]$openReportOnCompletion=$False,
			[Parameter(Mandatory=$true)][string]$itoContactName,
			[Parameter(Mandatory=$false)][string]$css,
			[Parameter(Mandatory=$true)]$xml)
{
	if ($xml.keys)
	{
		$csvfilesExported=@()
		$htmlPage = htmlHeader
		$htmlPage += "`n<h1>$reportHeader</h1>"
		$htmlPage += "`n<p>$reportIntro</p>"
		$reportTitles = $xml.keys | Where-Object {$_ -ne "Runtime"}
		logThis -msg "Keys = $([string]$reportTitles)"
		$index=1
		$reportTitles | ForEach-Object {
			#$subReportXML = $_
			$reportTitle = $_
			Write-Progress -Id 1 -Activity "Processing $reportTitle" -CurrentOperation "$index/$(($reportTitles | Measure-Object).Count)" -PercentComplete $($index/$(($reportTitles | Measure-Object).Count)*100)
			logThis -msg "[$index / $(($reportTitles | Measure-Object).Count)] reportTitle = $_"
			$htmlPage += "`n<h1>$($reportTitle)</h1>"

			#LogThis -msg "$($xml.$reportTitle.
			if ($xml.$reportTitle.Reports -and $xml.$reportTitle.Reports.keys)
			{
				$xml.$reportTitle.Reports.keys | ForEach-Object {
					$subReportsTitle = $_
					$htmlPage +=  "<h2>$subReportsTitle</h2>"
					#logThis -msg "$subReportsTitle Reports"
					$subReportsCount = ($xml.$reportTitle.Reports.$subReportsTitle.keys | measure-object).Count
					for ($jindex = 0 ; $jindex -lt $subReportsCount; $jindex++)
					{
						logThis -msg "$reportTitle\$subReportsTitle\$jindex"
						# Check if the subreport has a datatable. if NOT there is no point printing the section's other information.
						if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex])#.DataTable)
						{
							#logThis -msg "$reportTitle\$subReportsTitle\$jindex"
							if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].MetaData)
							{
								$metaData = convertTextVariablesIntoObject ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].MetaData)
								logThis -msg "`t`t-> [$($jindex+1)/$subReportsCount] $($metaData.tableHeader)"
							} else {
								showError -msg "`t`t-> No MetaData found in the object"
								logThis -msg "`t`t-> [$($jindex+1)/$subReportsCount] - Unknown"
							}
							if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].DataTable)
							{
								$reportTables = $xml.$reportTitle.Reports.$subReportsTitle[$jindex].DataTable

								$expectedPassesToMakeForTopConsumers = 2

								if ($metaData.generateTopConsumers)
								{
									# Report specifies that the report should generate a Top 10 (or other NUMBER) for this report, so need to make a second pass for the Top consumer report
									$passesToMake = $expectedPassesToMakeForTopConsumers
								} else {
									$passesToMake = 1
								}

								$passesSoFar=1;
								# If the report NFO file
								while ($passesSoFar -le $passesToMake)
								{
									if ($passesSoFar -eq $expectedPassesToMakeForTopConsumers)
									{
										$extraText = "(Top ($metaData.generateTopConsumers))"
									} else {
										$extraText = ""
									}

									if ($metaData.titleHeaderType)
									{
										logThis -msg "`t`t-> Using header Type $($metaData.titleHeaderType) found in NFO $extraText" # -ColourScheme $color
										#$headerType = $metaData.titleHeaderType # + " " + $extraText
										$headerType = $metaData.titleHeaderType#"h3"
									} else
									{
										logThis -msg "`t`t-> No header types found in NFO, using H2 instead $extraText"  -ColourScheme $global:colours.Information
										$headerType = "h1"
									}

									if ($metaData.tableHeader)
									{
										logThis -msg "`t`t-> Using heading from NFO: $($metaData.tableHeader) $extraText"  #-ColourScheme $color
										$htmlPage += "`n<$headerType>$($metaData.tableHeader) $extraText</$headerType>"
										#Remove-Variable tableHeader

									} else {
										logThis -msg "`t`t-> Will derive Heading from the original filename" #-ColourScheme $color
										if ($vcenterName)
										{
											#$title = $metaData.Filename.Replace($runtime+"_","").Replace("$vcenterName","").Replace(".csv","").Replace("_"," ")
											$title = $metaData.Filename.Replace("$vcenterName","").Replace(".csv","").Replace("_"," ") + " " + $extraText
										} else {
											#$title = $metaData.Filename.Replace($runtime+"_","").Replace(".csv","").Replace("_"," ")
											$title = $metaData.Filename.Replace(".csv","").Replace("_"," ") + " " + $extraText
										}
										$htmlPage += "`n<$headerType>$title $extraText</$headerType>"
										logThis -msg "`t`t-> $title $extraText" -ColourScheme $global:colours.Information
									}

									if ($metaData.introduction)
									{
										logThis -msg "`t`t-> Special introduction found in the NFO $extraText"  -ColourScheme $global:colours.Change
										#$htmlPage += "`n<p>$($metaData.introduction) $($metaData.analytics)</p>"
										$htmlPage += "`n<p>$($metaData.introduction)</p>"
										#Remove-Variable introduction
									} else {
										logThis -msg "`t`t-> No introduction found in the NFO $extraText"  -ColourScheme $global:colours.Information
									}
									if ($metaData.metaAnalytics)
									{
										#$htmlPage +=
									}

									$style="class=$setTableStyle"
									# sometime the dataTable in those reports can be a straigh forward Table (custom object) but at times, particularly
									# for the Performance results, it can be a hashtable with tables below for each metric. For this reason
									# i need to create a hashtable so I can feed it into the next phase.
									if ($reportTables -and ($reportTables.gettype().Name -eq "HashTable"))
									{
										#logThis -msg "This is HashTable"
										$tablesForProcessing = $reportTables
									} else {
										# because the dataTable for single tables is NOT a hashtable, I need to create one
										$tablesForProcessing = @{dataTable=$reportTables}
									}

									$tablesForProcessing.GetEnumerator() | ForEach-Object {
										$dataTable = $_.Value
										if($dataTable)
										{
											if ($passesSoFar -eq $expectedPassesToMakeForTopConsumers)
											{
												if ($metaData.generateTopConsumersSortByColumn)
												{
													$tmpreport = $dataTable | Sort-Object -Property $metaData.generateTopConsumersSortByColumn -Descending | Select-Object -First $metaData.generateTopConsumers
													$dataTable = $tmpreport
												} else {
													$tmpreport = $dataTable | Sort-object | Select-Object -First $metaData.generateTopConsumers
													$dataTable = $tmpreport
												}
											}
											if ($metaData.chartable -eq "true")
											{

												if ((Test-Path -path $metaData.ImageDirectory) -ne $true) {
													New-Item -type directory -Path $metaData.ImageDirectory
												}
												logThis -msg "-> Need to Chart this table according to NFO" -ColourScheme $global:colours.Change
												# do the charting here instead of the table
												$htmlPage += "`n<table>"
												#$chartStandardWidth
												#$chartStandardHeight
												#$imageFileType
												logThis -msg $dataTable -ColourScheme $global:colours.Information
												$report["OutputChartFile"] = createChart -sourceCSV $metaData.File -outputFileLocation $(($metaData.ImageDirectory)+$PATHSEPARATOR+$metaData.Filename.Replace(".csv",".$chartImageFileType")) -chartTitle $chartTitle `
													-xAxisTitle $xAxisTitle -yAxisTitle $yAxisTitle -imageFileType $chartImageFileType -chartType $chartType `
													-width $chartStandardWidth -height $chartStandardHeight -startChartingFromColumnIndex $startChartingFromColumnIndex -yAxisInterval $yAxisInterval `
													-yAxisIndex  $yAxisIndex -xAxisIndex $xAxisIndex -xAxisInterval $xAxisInterval

												$report["outputChartFileName"] = Split-Path -Leaf $metaData.OutputChartFile

												logThis -msg "-> image: ($metaData.OutputChartFile)" -ColourScheme $global:colours.Change
												$htmlPage += "`n<tr><td>"
												if ($emailReport)
												{
													$htmlPage += "`n<div><img src=""$($metaData.OutputChartFile)""></img></div>"
													$attachments += ($metaData.OutputChartFile)
												} else
												{
													$htmlPage += "`n<div><img src=""$($metaData.OutputChartFile)""></img></div>"
													# "<div><img src=""$($metaData.OutputChartFile)""></img></div>"
												}
												#$htmlPage += $metaData.DataTableCSV | ConvertTo-HTML -Fragment
												#$htmlPage += "`n</td></tr></table>"
												$htmlPage += "`n</td></tr>"
												#logThis -msg $imageFileLocation -ColourScheme $global:colours.Error
												#$attachments += $imageFileLocation
												#$htmlPage += "`n<div><img src=""$outputChartFileName""></img></div>"
												#$attachments += $imageFileLocation

												#Remove-Variable imageFileLocation
												#Remove-Variable imageFilename

												$htmlPage += "`n</td></tr></table>"
												#Remove-Variable chartable
											} else {
												# displayTableOrientation can be List or Table, must be set in the NFO file
												if (($dataTable | measure-Object).Count)
												{
													$count = ($dataTable | measure-Object).Count
												} else {
													$count = 1
												}
												$caption = ""
												#if (!$metaData.showTableCaption -or ($report.showTableCaption -eq $true))
												if ($metaData.showTableCaption -eq $true)
												{
													$caption = "<p>Number of items: $count</p>"

												}
												if ($metaData.displayTotals)
												{
													# add a last row in this Table and calculate the totals
													$tempTable = $dataTable
													#
													#
													#
													#
												}
												if ($metaData.displayTableOrientation -eq "List")
												{
													$dataTable | ForEach-Object {
														$htmlPage += ($_ | ConvertTo-HTML -Fragment -As $metaData.displayTableOrientation) -replace "<table","$caption<table class=aITTablesytle" -replace "&lt;/li&gt;","</li>" -replace "&lt\;li&gt;","<li>" -replace "&lt\;/ul&gt;","</ul>" -replace "&lt\;ul&gt;","<ul>"   -replace "`r","<br>"

													}
												} elseif ($metaData.displayTableOrientation -eq "Table") {
													#User has specified TABLE
													$htmlPage += ($dataTable | ConvertTo-HTML -Fragment -As "Table") -replace "<table","$caption<table class=aITTablesytle" -replace "&lt;/li&gt;","</li>" -replace "&lt\;li&gt;","<li>" -replace "&lt\;/ul&gt;","</ul>" -replace "&lt\;ul&gt;","<ul>"  -replace "`r","<br>"
													$htmlPage += "`n"

												} elseif ($metaData.displayTableOrientation -eq "TableTransposed") {
													#User has specified TABLE
													$htmlPage += ($dataTable | ConvertTo-HTML -Fragment -As "Table") -replace "<table","$caption<table class=aITTablesytle" -replace "&lt;/li&gt;","</li>" -replace "&lt\;li&gt;","<li>" -replace "&lt\;/ul&gt;","</ul>" -replace "&lt\;ul&gt;","<ul>"  -replace "`r","<br>"
													$htmlPage += "`n"

												} else {
													# Do this regardless
													$htmlPage += ($dataTable | ConvertTo-HTML -Fragment -As "Table") -replace "<table","$caption<table class=aITTablesytle" -replace "&lt;/li&gt;","</li>" -replace "&lt\;li&gt;","<li>" -replace "&lt\;/ul&gt;","</ul>" -replace "&lt\;ul&gt;","<ul>"  -replace "`r","<br>"
													$htmlPage += "`n"
													#$htmlPage += "`n<p>Invalid property in variable `$metaData.displayTableOrientation found in NFO. The only options are ""List"" and ""Table"" or don't put a variable in the NFO. <p>"
												}
												#Remove-Variable $report
											}
										} else {
											#$htmlPage += "`n<p><i>No items found.</i></p>"
										}
									}
									$passesSoFar++
								}
							} else {
								showError -msg "`t`t >>> No DataTable in the object to process <<<"
							} #	 if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex].DataTable)
						} else {
							showError -msg "No `$xml.$($reportTitle).Reports.$($subReportsTitle)[$($jindex)]"
						} #if ($xml.$reportTitle.Reports.$subReportsTitle[$jindex])
					} #for ($jindex = 0 ; $jindex -lt ($xml.$reportTitle.Reports.$subReportsTitle | Measure-Object).Count; $jindex++)
				} #$xml.$reportTitle.Reports.keys | ForEach-Object
			} else {
				showError -msg "No valid reports in `$xml.[category].Reports.keys"
			} # end if ($xml.$reportTitle.Reports.keys)
			$index++
		}
		$htmlPage += "`n<p>If you need clarification or require assistance, please contact $itoContactName ($replyToRecipients)</p><p>Regards,</p><p>$itoContactName</p>"
		$htmlPage += "`n"
		$htmlPage += htmlFooter
		return $htmlPage
	} else {
		logThis -msg "Invalid input `$xml.keys file"
	}
}

function returnText ([Parameter(Mandatory=$true)][string]$txt)
{
	if ($txt)
	{
		return $txt
	} else {
		return "N/A"
	}
}

function convertTextVariablesIntoObject ([Parameter(Mandatory=$true)][object]$obj,[Parameter(Mandatory=$false)][object]$makeGlobalVars=$false)
{
	#logThis -msg "Reading in configurations from file $inifile"
	$configurations = @{}
	$obj | ForEach-Object {
		if ($_ -notlike "#*")
		{
			$var = $_.Split('=')
			#logThis -msg $var
			#logThis -msg $var[0]
			if (($var | measure-Object).Count -gt 1)
			{
				$name=$var[0]
				#logThis -msg "$($var[0]) $($var[1])"
				if ($var[1] -eq "true")
				{
					#$configurations.Add($var[0],$true)
					$value=$true
					#New-Variable -Name $var[0] -Value $true
				} elseif ($var[1] -eq "false")
				{
					#$configurations.Add($var[0],$false)
					$value=$false
					#New-Variable -Name $var[0] -Value $false
				} else {
					if ($var[1] -match ',')
					{
						$value = $var[1] -split ','
						#New-Variable -Name $var[0] -Value ($var[1] -split ',')
					} else {
						$value = $var[1]
						#New-Variable -Name $var[0] -Value $var[1]
					}
				}
				$configurations.Add($name,$value)
				if ($makeGlobalVars)
				{
					Set-Variable -Scope Global -Name $name -Value $value
				}
			}
		}
	}


	if ($configurations)
	{
		# Perform post processing by replace all strings with  $ sign in them with the content of their respective Content.
		# for example: replaceing $customer with the actual customer name specified by the key $configurations.customer
		$postProcessConfigs = @{}
		$configurations.Keys | ForEach-Object {
			$keyname=$_
			#logThis -msg $keyname
			# just in case the value is an array, process each
			$updatedValue=""
			#logThis -msg  $configurations.$keyname		 -ColourScheme White
			$updatedField = $configurations.$keyname | ForEach-Object {
				$curr_string = $_

				if ($curr_string -match '\$')
				{
					# replace the string with a $ sign in it with the content of the variable it is expected
					$newstring=""
					$newstring_array = $curr_string -split ' ' | ForEach-Object {
						$word = $_
						#logThis -msg "`tBefore $word"
						if ($word -like '*$*')
						{
							$key=$word -replace '\$'
							$configurations.$key
							#logThis -msg "Needs replacing $word with $($configurations.$key)"
						} else {
							$word
							#logThis -msg "$($word)"
						}
					}
					$updatedValue = [string]$newstring_array
					#logThis -msg "-t>>>$updatedValue" -ColourScheme $global:colours.Change
				} elseif ($curr_string -eq $true)
				{
					$updatedValue = $true
					#logThis -msg "-t>>>$updatedValue" -ColourScheme $global:colours.Information
				} elseif ($curr_string -eq $false)
				{
					$updatedValue = $false
					#logThis -msg "-t>>>$updatedValue" -ColourScheme $global:colours.Highlight
				} else {

					$updatedValue = $curr_string
					#logThis -msg "-t>>>$updatedValue" -ColourScheme $global:colours.Information
				}
				$updatedValue
			}
			$postProcessConfigs.Add($keyname,$updatedField)
		}
		#$postProcessConfigs.Add("inifile",$inifile)
		#return $configurations,$postProcessConfigs

		return $postProcessConfigs#,$configurations
	} else {
		return $null
	}
}


Function Open-ExcelPackage  {
<#
.Synopsis
    Returns an Excel Package Object with for the specified XLSX ile
.Example
    $excel  = Open-ExcelPackage -path $xlPath
    $sheet1 = $excel.Workbook.Worksheets["sheet1"]
    set-Format -Address $sheet1.Cells["E1:S1048576"], $sheet1.Cells["V1:V1048576"]  -NFormat ([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern)
    close-ExcelPackage $excel -Show
   This will open the file at $xlPath, select sheet1 apply formatting to two blocks of the sheet and close the package
#>
    [OutputType([OfficeOpenXml.ExcelPackage])]
    Param ([Parameter(Mandatory=$true)]$Path,
           [switch]$KillExcel)

        if($KillExcel)         {
            Get-Process -Name "excel" -ErrorAction Ignore | Stop-Process
            while (Get-Process -Name "excel" -ErrorAction Ignore) {}
        }

        $Path          = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if (Test-Path $path) {New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path }
        Else                 {Write-Warning "Could not find $path" }
 }

Function Close-ExcelPackage {
<#
.Synopsis
    Closes an Excel Package, saving, saving under a new name or abandoning changes and opening the file as required
#>
    Param (
    #File to close
    [parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [OfficeOpenXml.ExcelPackage]$ExcelPackage,
    #Open the file
    [switch]$Show,
    #Abandon the file without saving
    [Switch]$NoSave,
    #Save file with a new name (ignored if -NoSaveSpecified)
    $SaveAs
    )
    if ( $NoSave)      {$ExcelPackage.Dispose()}
    else {
          if ($SaveAs) {$ExcelPackage.SaveAs( $SaveAs ) }
          Else         {$ExcelPackage.Save(); $SaveAs = $ExcelPackage.File.FullName }
          $ExcelPackage.Dispose()
          if ($show)   {Start-Process -FilePath $SaveAs }
    }
}


#######################################################################################
# MAIN
#if (!$global:logDir)
#{
	#Set-Variable -Name logDir -Value $MyInvocation.MyCommand.Path -Scope Global
#}
if ($global:report.Runtime.LogDirectory)
{
	$global:logDir = $global:report.Runtime.LogDirectory
}

if($global:logDir -and ((Test-Path -path $global:logDir) -ne $true)) {
	New-Item -type directory -Path $global:logDir
	$childitem = Get-Item -Path $global:logDir
	$global:logDir = $childitem.FullName
}

$global:runtime=getShortDateTime
$global:today = getShortDate
$global:startDate = (Get-Date (forThisdayGetFirstDayOfTheMonth -day (get-date $today)) -Format "dd-MM-yyyy midnight")
$global:lastDate = (Get-date ( forThisdayGetLastDayOfTheMonth -day $(Get-Date $today).AddMonths(-$showLastMonths) ) -Format "dd-MM-yyyy midnight")
$global:runtimeLogFileInMemory=""

if ($global:scriptName -and $global:logDir)
{
	#SetmyCSVOutputFile -filename "$($global:logDir)$($PATHSEPARATOR)$($($global:scriptName).Replace(".ps1",".csv"))"
	#SetmyCSVMetaFile -filename "$($global:logDir)$($PATHSEPARATOR)$($($global:scriptName).Replace(".ps1",".nfo"))"
	SetmyLogFile -filename "$($global:logDir)$($PATHSEPARATOR)$($($global:scriptName).Replace(".ps1",".log"))"
	#$global:runtimeLogFile = "$($global:logDir)$($PATHSEPARATOR)$($($global:scriptName).Replace(".ps1",".log"))"
} else {
	#SetmyCSVOutputFile -filename "$($global:logDir)$($PATHSEPARATOR)genericModule.csv"
	#SetmyCSVMetaFile -filename "$($global:logDir)$($PATHSEPARATOR)genericModule.nfo"
	SetmyLogFile -filename ".$($PATHSEPARATOR)genericModule.log"
	#$global:runtimeLogFile = ".$($PATHSEPARATOR)genericModule.log"
}
logThis -msg "****************************************************************************" -ColourScheme $global:colours.Highlight
logThis -msg "Script Started @ $(getdate)" -ColourScheme $global:colours.Highlight
logThis -msg "Executing script: $global:scriptName " -ColourScheme $global:colours.Highlight
logThis -msg "runtime: $($global:runtime)" -ColourScheme $global:colours.Highlight
logThis -msg "today: $($global:today)" -ColourScheme $global:colours.Highlight
logThis -msg "startDate: $($global:startDate)" -ColourScheme $global:colours.Highlight
logThis -msg "lastDate: $($global:lastDate)" -ColourScheme $global:colours.Highlight
logThis -msg "runtimeLogFileInMemory: $($global:runtimeLogFileInMemory)" -ColourScheme $global:colours.Highlight

#logThis -msg "Output Dir = $global:logDir" -ColourScheme $global:colours.Highlight
#logThis -msg " Runtime log file = $global:runtimeLogFile" -ColourScheme $global:colours.Highlight
#logThis -msg " Runtime CSV File = $global:runtimeCSVOutput" -ColourScheme $global:colours.Highlight
#logThis -msg " Runtime Meta File = $global:runtimeCSVMetaFile" -ColourScheme $global:colours.Highlight
#logThis -msg " Runtime Meta File (In Memory) = $global:runtimeLogFileInMemory" -ColourScheme $global:colours.Highlight
logThis -msg "****************************************************************************" -ColourScheme $global:colours.Highlight

<# OLD STUFF TO CLEAR OUT
#$global:runtime="$(date -f dd-MM-yyyy)"

#$childitem = Get-Item -Path $global:logDir
#$global:logDir = $childitem.FullName
#$runtimeLogFile = $global:logDir + +$runtime+"_"+$global:scriptName.Replace(".ps1",".log")
#$global:runtimeCSVOutput = $global:logDir + +$runtime+"_"+$global:scriptName.Replace(".ps1",".csv")
#$runtimeCSVMetaFile = $global:logDir + +$runtime+"_"+$global:scriptName.Replace(".ps1",".nfo")

$runtimeLogFile = $global:logDir + +$global:scriptName.Replace(".ps1",".log")
$global:runtimeCSVOutput = $global:logDir++$global:scriptName.Replace(".ps1",".csv")
$runtimeCSVMetaFile = $global:logDir++$global:scriptName.Replace(".ps1",".nfo")
$scriptsHomeDir = split-path -parent $global:scriptName

$global:today = Get-Date
$global:startDate = (Get-Date (forThisdayGetFirstDayOfTheMonth -day $today) -Format "dd-MM-yyyy midnight")
$global:lastDate = (Get-date ( forThisdayGetLastDayOfTheMonth -day $(Get-Date $today).AddMonths(-$showLastMonths) ) -Format "dd-MM-yyyy midnight")

if (!$global:reportIndex)
{

	logThis -msg "Creating the Indexer File $global:logDir\index.txt"
	setReportIndexer -fullName "$global:logDir\index.txt"
}

SetmyLogFile -filename $runtimeLogFile
logThis -msg " ****************************************************************************"
logThis -msg "Script Started @ $(get-date)" -ColourScheme $global:colours.Information
logThis -msg "Executing script: $global:scriptName " -ColourScheme $global:colours.Information
logThis -msg "Logging Directory: $global:logDir" -ColourScheme  $global:colours.Yellow
logThis -msg "Script Log file: $global:logfile" -ColourScheme  $global:colours.Yellow
logThis -msg "Indexer: $global:reportIndex" -ColourScheme  $global:colours.Yellow
logThis -msg " ****************************************************************************"
logThis -msg "Loading Session Snapins.."
#loadSessionSnapings
SetmyCSVOutputFile -filename $global:runtimeCSVOutput
SetmyCSVMetaFile -filename $runtimeCSVMetaFile
#>

# Sometime tables can be in returned in a Name/Value format
# Example source:
# 		Headers: Name			Value
# 		Row 1  : id     	383840-34-34
# 		Row 2  : Name   	Jason
# 		Row 3  : Lastname Bourne
# Example after conversion:
# 		Headers: Id							Name			Lastname
# 		Row 1  : 383840-34-34   Jason			Bourne

function transposeNameValueTableToArray ([Parameter(Mandatory=$true)][object]$keyValueTable)
{
	$table = New-Object System.Object
	$index=0
	$keyValueTable | ForEach-Object {
		$table | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value
	}
	if (($table | Measure-Object).Count -gt 0)
	{
		$table
	}
}