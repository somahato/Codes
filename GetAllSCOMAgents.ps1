$Cur = Get-Location
Remove-Item -path "$Cur\Agentlist.txt"
Import-Module OperationsManager
$Agentlist = Get-scomagent
foreach ($Agent in $Agentlist)

{
$DisplayName=$Agent.Displayname
$IPAddress=$Agent.IPAddress
$HealthState=$Agent.HealthState
$Domain = $Agent.Domain

"ServerName =""$DisplayName"",IPAddress=""$IPAddress"",Domain=""$Domain"",HealthState=""$HealthState""" >> "$Cur\Agentlist.txt"
}
