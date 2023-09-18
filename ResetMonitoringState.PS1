<###########################################################################################
.Author
Sourav Mahato
.Created Date
18/09/2023
.Purpose
Script for to Reset the health of the Monitor where alert was already closed.           
.How to run
Save the Script as .PS1 and run it from a SCOM Management Server.             
e.g PS C:\Users\sourav.LAB> .\GetAllSCOMAgents.ps1                                           
############################################################################################>

$Alerts = Get-SCOMAlert | Where-Object {$_.ResolutionState -ne "0"}
ForEach($Alert in $Alerts) {
    $Monitor = Get-SCOMMonitor -Id $Alert.MonitoringRuleId
    Get-SCOMClassInstance -ID $Alert.MonitoringObjectId | ForEach {$_.ResetMonitoringState($Monitor)}
}
