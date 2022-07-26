$cur = get-location

$Servers = Get-content "$cur\batch1.txt"


Write-Host "Checking if OperationsManager Module is loaded"            

#Check if OperationsManager Module is loaded.
                 
if(Get-Module OperationsManager)           

{            
 Write-Host "OperationsManager Module is already imported" -ForegroundColor Yellow          
}

else
{

Write-Host "Importing OperationsManager Module" -ForegroundColor Yellow

Import-Module OperationsManager -ErrorAction Stop            

}

$MonitoringObject = Get-SCOMClass -Name "Microsoft.Windows.Computer" | Get-SCOMClassInstance #| ? {($_.DisplayName -notlike 'SCOMMS12019*' -And $_.DisplayName -notlike 'SCOMMS22019*')}


foreach($Server in $Servers)

{

    $MSValidation = Get-SCOMManagementServer | where {$_.displayname -like "$Server"}


    if (!$MSValidation)
    {

        $Object = $MonitoringObject | where {$_.displayname -like "$Server"}

        $ti = New-Timespan -Hour 5 -Minute 15

        $Time = ((Get-Date).Add($ti))

        if($Object.InMaintenanceMode -eq $false)

        {

          "Putting " + $Object.Displayname + " into maintenance mode"

           Start-SCOMMaintenanceMode -Instance $Object -EndTime:$Time -Reason: "PlannedOther" -Comment: "scheduled Activity - daily Reboot"
        }

        else 

        {

            "Server " + $Object.Displayname + " is already into maintenance mode"
 
        }
    }

    Else
    
    {
   
    Write-Host "Server $Server is a Management server" -ForegroundColor Yellow

    "Server $Server is a Management server" >> "$cur\MMLogs.txt"
   
    }
 
 }  