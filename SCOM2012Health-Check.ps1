##############################################################################
#
#   SCOM2007R2Health-Check.ps1
#   Created By Sourav Mahato
#   Created on 16th October 2015
##############################################################################

if (Test-path "C:\Scripts\HealthCheck.html")

{
  Remove-Item -Path "C:\Scripts\HealthCheck.html" -Force
}

#Importing the SCOM PowerShell module

Import-module OperationsManager

#Connect to localhost when running on the management server

$connect = New-SCOMManagementGroupConnection –ComputerName localhost

# Create header for HTML Report

$Head = "<style>"
$Head +="BODY{background-color:#CCCCCC;font-family:Verdana,sans-serif; font-size: x-small;}"
$Head +="TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; width: 100%;}"
$Head +="TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:green;color:white;padding: 5px; font-weight: bold;text-align:left;}"
$Head +="TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#F0F0F0; padding: 2px;}"
$Head +="</style>"

# Get status of Management Server Health and input them into report

write-host "Getting SCOM Management Servers Health Status" -ForegroundColor Yellow 
$ReportOutput = "To enable HTML view, click on `"This message was converted to plain text.`" and select `"Display as HTML`""

$MonitoringMS = Get-SCOMClass -Name "Microsoft.SystemCenter.ManagementServer" | Get-SCOMClassInstance

$ReportOutput += "<p><H2>All Management Servers Health Status in SCOM console</H2></p>"

$Count = $MonitoringMS | where {($_.IsAvailable -eq $False -and $_.InMaintenanceMode -eq $False)} | Measure-Object

if($Count.Count -gt 0)
{ 
 $ReportOutput += $MonitoringMS | where {($_.IsAvailable -eq $False -and $_.InMaintenanceMode -eq $False)} | select DisplayName,HealthState,InMaintenanceMode |ConvertTo-HTML -fragment
}
else
{ 
 $ReportOutput += "<p>All SCOM Management Servers are in Healthy State in SCOM Console.</p>"
} 

##############################################################################

# Get status of Maintenance Mode for Management Server

write-host "Getting SCOM Servers Maintenance Mode Status" -ForegroundColor Yellow

$ReportOutput += "<p><H2>All SCOM Management Servers Maintenance Mode Status</H2></p>"

$Count = $MonitoringMS | where {$_.InMaintenanceMode -eq $True} | Measure-Object

if($Count.Count -gt 0)
{ 

$ReportOutput += $MonitoringMS | where {$_.InMaintenanceMode -eq $True} | select DisplayName,HealthState,InMaintenanceMode |ConvertTo-HTML -fragment

}

Else
{
$ReportOutput += "<p>SCOM Management Servers are not in maintenance Mode.</p>"

}

##############################################################################

# Get SCOM Agent Servers Which are in Maintenance Mode into report

write-host "Getting SCOM Agent Servers Which are in Maintenance Mode" -ForegroundColor Yellow

$ReportOutput += "<p><H2>SCOM Agent Servers Which are in Maintenance Mode</H2></p>"
    
$criteria = new-object Microsoft.EnterpriseManagement.Monitoring.MonitoringObjectGenericCriteria("InMaintenanceMode=1")
$ManagementGroup = Get-SCOMManagementGroup
$objectsInMM = $ManagementGroup.GetPartialMonitoringObjects($criteria)
$ObjectsFound = $objectsInMM | select-object DisplayName, @{name="Object Type";expression={foreach-object {$_.GetLeastDerivedNonAbstractMonitoringClass().DisplayName}}},@{name="StartTime";expression={foreach-object {$_.GetMaintenanceWindow().StartTime.ToLocalTime()}}},@{name="EndTime";expression={foreach-object {$_.GetMaintenanceWindow().ScheduledEndTime.ToLocalTime()}}},@{name="Path";expression={foreach-object {$_.Path}}},@{name="User";expression={foreach-object {$_.GetMaintenanceWindow().User}}},@{name="Reason";expression={foreach-object {$_.GetMaintenanceWindow().Reason}}},@{name="Comment";expression={foreach-object {$_.GetMaintenanceWindow().Comment}}}

$MMObject = Get-SCOMClass -Name "Microsoft.Windows.Computer" | Get-SCOMClassInstance | where {$_.InMaintenanceMode -eq $True}
$Count = $MMObject | Measure-Object

if($Count.Count -gt 0)
{ 
$Agents = $MMObject | Sort-Object HealthState -descending | select DisplayName,HealthState

$AgentTable = New-Object System.Data.DataTable "$AvailableTable"
$AgentTable.Columns.Add((New-Object System.Data.DataColumn DisplayName,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn HealthState,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MM,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMUser,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMReason,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMComment,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMStartTime,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMEndTime,([string])))

foreach ($Agent in $Agents)
   {
    $FoundObject = $null
    $MaintenanceModeUser = $null
    $MaintenanceModeComment = $null
    $MaintenanceModeReason = $null
    $MaintenanceModeStartTime = $null
    $MaintenanceModeEndTime = $null
    $FoundObject = 0
    $FoundObject = $objectsFound | ? {$_.DisplayName -match $Agent.DisplayName}
   
     if ($FoundObject -ne $null)
        {
         $MaintenanceMode = "Yes"
         $MaintenanceObjectCount = $FoundObject.Count
         $MaintenanceModeUser = (($FoundObject | Select User)[0]).User
         $MaintenanceModeReason = (($FoundObject | Select Reason)[0]).Reason
         $MaintenanceModeComment = (($FoundObject | Select Comment)[0]).Comment
         $MaintenanceModeStartTime = ((($FoundObject | Select StartTime)[0]).StartTime).ToString()
         $MaintenanceModeEndTime = ((($FoundObject | Select EndTime)[0]).EndTime).ToString()
        }

        $NewRow = $AgentTable.NewRow()
        $NewRow.DisplayName = ($Agent.DisplayName).ToString()
        $NewRow.HealthState = ($Agent.HealthState).ToString()
        $NewRow.MM = $MaintenanceMode
        $NewRow.MMUser = $MaintenanceModeUser
        $NewRow.MMReason = $MaintenanceModeReason
        $NewRow.MMComment = $MaintenanceModeComment
        $NewRow.MMStartTime = $MaintenanceModeStartTime
        $NewRow.MMEndTime = $MaintenanceModeEndTime
        $AgentTable.Rows.Add($NewRow)
    }
    
$ReportOutput += $AgentTable | Sort-Object MMEndTime | Select DisplayName, HealthState, MM, MMUser, MMReason, MMComment, MMStartTime, MMEndTime | ConvertTo-HTML -fragment
}

else

{ 

 $ReportOutput += "<p>SCOM Managed Servers are not in maintenance Mode. All are in Monitored State.</p>"

} 

###################################################################################################################

# Get Clusters Which are in Maintenance Mode into report

write-host "Getting Clusters Which are in Maintenance Mode" -ForegroundColor Yellow

$ReportOutput += "<p><H2>Clusters Which are in Maintenance Mode</H2></p>"
    
$criteria = new-object Microsoft.EnterpriseManagement.Monitoring.MonitoringObjectGenericCriteria("InMaintenanceMode=1")
$ManagementGroup = Get-SCOMManagementGroup
$objectsInMM = $ManagementGroup.GetPartialMonitoringObjects($criteria)
$ObjectsFound = $objectsInMM | select-object DisplayName, @{name="Object Type";expression={foreach-object {$_.GetLeastDerivedNonAbstractMonitoringClass().DisplayName}}},@{name="StartTime";expression={foreach-object {$_.GetMaintenanceWindow().StartTime.ToLocalTime()}}},@{name="EndTime";expression={foreach-object {$_.GetMaintenanceWindow().ScheduledEndTime.ToLocalTime()}}},@{name="Path";expression={foreach-object {$_.Path}}},@{name="User";expression={foreach-object {$_.GetMaintenanceWindow().User}}},@{name="Reason";expression={foreach-object {$_.GetMaintenanceWindow().Reason}}},@{name="Comment";expression={foreach-object {$_.GetMaintenanceWindow().Comment}}}

$MMObject1 = Get-SCOMClass -Name "Microsoft.Windows.Cluster" | Get-SCOMClassInstance | where {$_.InMaintenanceMode -eq $True}
$Count1 = $MMObject1 | Measure-Object

if($Count1.Count -gt 0)
{ 
$Agents = $MMObject1 | Sort-Object HealthState -descending | select DisplayName,HealthState

$AgentTable = New-Object System.Data.DataTable "$AvailableTable"
$AgentTable.Columns.Add((New-Object System.Data.DataColumn DisplayName,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn HealthState,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MM,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMUser,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMReason,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMComment,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMStartTime,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMEndTime,([string])))

foreach ($Agent in $Agents)
   {
    $FoundObject = $null
    $MaintenanceModeUser = $null
    $MaintenanceModeComment = $null
    $MaintenanceModeReason = $null
    $MaintenanceModeStartTime = $null
    $MaintenanceModeEndTime = $null
    $FoundObject = 0
    $FoundObject = $objectsFound | ? {$_.DisplayName -match $Agent.DisplayName}
   
     if ($FoundObject -ne $null)
        {
         $MaintenanceMode = "Yes"
         $MaintenanceObjectCount = $FoundObject.Count
         $MaintenanceModeUser = (($FoundObject | Select User)[0]).User
         $MaintenanceModeReason = (($FoundObject | Select Reason)[0]).Reason
         $MaintenanceModeComment = (($FoundObject | Select Comment)[0]).Comment
         $MaintenanceModeStartTime = ((($FoundObject | Select StartTime)[0]).StartTime).ToString()
         $MaintenanceModeEndTime = ((($FoundObject | Select EndTime)[0]).EndTime).ToString()
        }

        $NewRow = $AgentTable.NewRow()
        $NewRow.DisplayName = ($Agent.DisplayName).ToString()
        $NewRow.HealthState = ($Agent.HealthState).ToString()
        $NewRow.MM = $MaintenanceMode
        $NewRow.MMUser = $MaintenanceModeUser
        $NewRow.MMReason = $MaintenanceModeReason
        $NewRow.MMComment = $MaintenanceModeComment
        $NewRow.MMStartTime = $MaintenanceModeStartTime
        $NewRow.MMEndTime = $MaintenanceModeEndTime
        $AgentTable.Rows.Add($NewRow)
    }
    
$ReportOutput += $AgentTable | Sort-Object MMEndTime | Select DisplayName, HealthState, MM, MMUser, MMReason, MMComment, MMStartTime, MMEndTime | ConvertTo-HTML -fragment
}

else

{ 

 $ReportOutput += "<p>No Cluster is in maintenance Mode.</p>"

}

##############################################################################################
# Get status of SCOM Agent Servers Which are in Grayed Out State and input them into report

write-host "Getting SCOM Agent Servers Which are in Grayed Out State" -ForegroundColor Yellow

$MonitoringObject = Get-SCOMClass -Name "Microsoft.SystemCenter.Agent" | Get-SCOMClassInstance

$ReportOutput += "<p><H2>SCOM Agent Servers Which Are in Grayed Out State</H2></p>"

$AgentCount = $MonitoringObject | where {($_.IsAvailable -eq $False -and $_.InMaintenanceMode -eq $False)} | Measure-Object

if($AgentCount.Count -gt 0)
{ 
 $ReportOutput += $MonitoringObject | where {($_.IsAvailable -eq $False -and $_.InMaintenanceMode -eq $False)} | select DisplayName,HealthState,InMaintenanceMode |ConvertTo-HTML -fragment
}
else
{ 
 $ReportOutput += "<p>All SCOM Agent Servers are in Healthy State in SCOM Console.</p>"
 
} 

####################################################################################################

# Get status of SCOM Agent Servers Which are in Not Monitored State and input them into report

write-host "Getting SCOM Agent Servers Which are in Not Monitored State" -ForegroundColor Yellow

$ReportOutput += "<p><H2>SCOM Agent Servers Which Are in Not Monitored State</H2></p>"

$AgentCount1 = $MonitoringObject | where {($_.IsAvailable -eq $True -and $_.InMaintenanceMode -eq $False -and $_.HealthState -eq "Uninitialized")} | Measure-Object

if($AgentCount1.Count -gt 0)
{ 
 $ReportOutput += $MonitoringObject | where {($_.IsAvailable -eq $True -and $_.InMaintenanceMode -eq $False -and $_.HealthState -eq "Uninitialized")} | select DisplayName,HealthState,InMaintenanceMode |ConvertTo-HTML -fragment
}
else
{ 
 $ReportOutput += "<p>There are no SCOM Agent Server is in Not Monitored State in SCOM Console.</p>"
} 

##############################################################################

# Get Alerts specific to Management Servers and put them in the report
write-host "Getting Management Server Alerts" -ForegroundColor Yellow 
$ReportOutput += "<h2>Management Server Alerts</h2>"
$ManagementServers = Get-SCOMManagementServer

foreach ($ManagementServer in $ManagementServers)
{ 
 $ReportOutput += "<h3>Alerts on " + $ManagementServer.ComputerName + "</h3>"

$MS = $ManagementServer.Name

 $MSAlerts= get-SCOMalert -Criteria ("NetbiosComputerName = '" + $ManagementServer.ComputerName + "'") | where {$_.ResolutionState -ne '255' -and $_.MonitoringObjectFullName -Match 'Microsoft.SystemCenter' -and $_.severity -Match 'Error'} | Measure-Object

 if($MSAlerts.Count -gt 0)

 {
 
    $ReportOutput += get-SCOMalert -Criteria ("NetbiosComputerName = '" + $ManagementServer.ComputerName + "'") | where {$_.ResolutionState -ne '255' -and $_.MonitoringObjectFullName -Match 'Microsoft.SystemCenter'} | select TimeRaised,Name,Description,Severity | ConvertTo-HTML -fragment
 
 }
 
 Else

 {

    $ReportOutput += "<p>There are no Critical Alerts Present for The $MS in SCOM Console.</p>"
 
 }

}

##############################################################################

# Get all alerts

write-host "Getting all alerts" -ForegroundColor Yellow

$Alerts = Get-SCOMAlert | where {$_.ResolutionState -ne '255'}

$AllAlerts = Get-SCOMAlert | where {$_.severity -Match 'Error'}

##############################################################################

# Get alerts for last 24 hours

write-host "Getting Critical alerts for last 24 hours" -ForegroundColor Yellow

$ReportOutput += "<h2>Top 20 Critical Alerts With Same Name - 24 hours</h2>"
$ReportOutput += $AllAlerts | where {$_.LastModified -le (Get-Date).addhours(-24)} | Group-Object Name | Sort-object Count -desc | select-Object -first 20 Count, Name, ResolutionState | ConvertTo-HTML -fragment

$ReportOutput += "<h2>Top 20 Repeating Active Critical Alerts Last Modified in 24 hours</h2>"
$ReportOutput += $Alerts | where {$_.LastModified -le (Get-Date).addhours(-24)} | Sort-Object -desc RepeatCount | select-Object -first 20 RepeatCount,ResolutionState, Name, MonitoringObjectPath, Description | ConvertTo-HTML -fragment

##############################################################################

# Get the Top 10 Unresolved alerts still in console and put them into report

write-host "Getting Top 10 Unresolved Critical Alerts With Same Name - All Time" -ForegroundColor Yellow 

$ReportOutput += "<h2>Top 10 Unresolved Critical Alerts With Same Name - All Time</h2>"
$ReportOutput += $Alerts  | Group-Object Name | Sort-object Count -desc | select-Object -first 10 Count, Name | ConvertTo-HTML -fragment

##############################################################################

# Get the Top 10 Repeating Alerts and put them into report

write-host "Getting Top 10 Repeating Critical Alerts - All Time" -ForegroundColor Yellow 

$ReportOutput += "<h2>Top 10 Repeating Critical Alerts - All Time</h2>"
$ReportOutput += $AllAlerts | Sort -desc RepeatCount | select-object –first 10 Name, RepeatCount, MonitoringObjectPath, Description, ResolutionState | ConvertTo-HTML -fragment

##############################################################################

# Get list of agents still in Pending State and put them into report

write-host "Getting Agents in Pending State" -ForegroundColor Yellow 

$ReportOutput += "<h2>Getting Agents in Pending State</h2>"

$AgentPendingAction = Get-SCOMPendingManagement | Measure-Object

if($AgentPendingAction.Count -gt 0)
{ 
 $ReportOutput += Get-SCOMPendingManagement | sort AgentPendingActionType | select AgentName,ManagementServerName,AgentPendingActionType | ConvertTo-HTML -fragment
}
else
{ 
 $ReportOutput += "<p>No New SCOM Agent Server is in Pending Management Console as a Pending State.</p>"
} 

##############################################################################

# List Management Packs updated in last 24 hours

write-host "Getting List Management Packs updated in last 24 hours" -ForegroundColor Yellow

$ReportOutput += "<h2>Management Packs Updated</h2>"

$MPDates = (Get-Date).adddays(-1)

$MPs = Get-SCManagementPack | Where {$_.LastModified -gt $MPDates} | Measure-Object

if($MPs.Count -gt 0)
{ 
 $ReportOutput += Get-SCManagementPack | Where {$_.LastModified -gt $MPDates} | Select-Object DisplayName, LastModified | Sort LastModified | ConvertTo-Html -fragment
}
else
{ 
 $ReportOutput += "<p>There are no Management Pack got Updated in Last 24 hours.</p>"

}

##############################################################################

# Take all $ReportOutput and combine it with $Body to create completed HTML output

$Body = ConvertTo-HTML -head $Head -body "$ReportOutput"

##############################################################################

#Send email functionality from below line, use it if you want   
 
$smtpServer = "exbhsrv.internal.lr.org"
$Body = ConvertTo-HTML -head $Head -body "$ReportOutput"
$messageSubject = "SCOM 2012 Daily Healthcheck Report" 
$message = New-Object System.Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$message.from = "SVC-OMSvrActionAccou@lr.org"
$message.To.Add("sourav.mahato@capgemini.com")
#$message.To.Add("sohini.moitra@capgemini.com")
$message.To.Add("james.x.gordon@capgemini.com")
$message.CC.Add("khaja.baig@capgemini.com")
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = $Body
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)