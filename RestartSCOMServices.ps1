Import-Module OperationsManager

$ManagementServers = Get-SCOMManagementServer | ? {$_.IsGateway -eq $False}

$MSs = $ManagementServers.Displayname

Foreach ($MS in $MSs)

{
Invoke-Command -ComputerName $MS {restart-service  omsdk, cshost, healthService}
}
