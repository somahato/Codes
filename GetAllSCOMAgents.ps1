$Cur = Get-Location
if (Test-Path "$Cur\Agentlist.txt")
        {
            Remove-Item -Path "$Cur\Agentlist.txt" -Force
        }

if (Test-Path "$Cur\UnixAgentlist.txt")
        {
            Remove-Item -Path "$Cur\UnixAgentlist.txt" -Force
        }

Import-Module OperationsManager

$Agentlist = Get-SCOMAgent

foreach ($Agent in $Agentlist)
{

$DisplayName=$Agent.Displayname
$IPAddress=$Agent.IPAddress
$HealthState=$Agent.HealthState
$Domain = $Agent.Domain

"DisplayName=""$DisplayName"",IPAddress=""$IPAddress"",HealthState=""$HealthState""" >> $Cur\Agentlist.txt
}

$SCXAgentlist = Get-SCOMClass -Displayname “UNIX/Linux Computer” | Get-SCOMClassInstance

foreach ($Agent in $SCXAgentlist)
{

$DisplayName=$Agent.Displayname
$IPAddress=$Agent.'[Microsoft.Unix.Computer].IPAddress'
$HealthState=$Agent.HealthState
$Domain = $Agent.Domain

"DisplayName =""$DisplayName"",IPAddress=""$IPAddress"",HealthState=""$HealthState""" >> $Cur\Agentlist.txt
}
