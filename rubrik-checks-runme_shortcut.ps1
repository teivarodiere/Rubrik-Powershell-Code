<# 
rubrik-checks-run-me.ps1 
#>

$xmloutput = .\rubrik-checks.ps1 -Passwordfile "D:\scripts\rubrik\rubrikcreds.txt" -username "DOMAIN\User" -Target "x.x.x.x" -farmName "Customer Farm" -itoContactName "My IT Company" -reportPeriodDays 90
return $xmloutput