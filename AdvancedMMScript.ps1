Import-Module operationsmanager

$file = "C:\Softwares\Servers.txt"

$endtime = (get-date).AddMinutes(10) #Time duration for Maintenance Mode in minutes#
$class = get-scomclass | where {$_.Displayname -like "Windows Computer"} #selecting Windows computer class#

$instance = Get-SCOMClassInstance -Class $class

$line = Get-Content -Path $file | Measure-Object -Line

If($line)
{
    $decimal = $line.Lines/100
    #$decimal
    $BatchCount = [math]::round($decimal)
    #$BatchCount

For ($i=1; $i -le $BatchCount; $i++) 

    {
    
    $i

      Start-Sleep -Seconds 120

    Write-Host "Need to Skip the objects from Batch $i" -ForegroundColor Cyan

     If ($i -eq 1)
        {

            $Batch = Get-Content -Path $file | select -First 100

            $Batch >> "C:\Softwares\$i.txt"

            Write-Host "Going to create first batch for object $No"

            foreach ($s in $Batch)

            {
            $server = $instance | where {($_.DisplayName -like $s) -or ($_.DisplayName -contains $s+".contoso.local") } #selecting the windows computer object of each server#

                if($server.InMaintenanceMode -like "False") #Neglecting servers which are already in MM#

                {
                Start-SCOMMaintenanceMode -Reason PlannedApplicationMaintenance -Comment "Planned Maintenance" -Instance $server -EndTime $endtime

                write-host $server "has been kept in MM till $endtime" #display output#
                }

                Elseif($server.InMaintenanceMode -like "True")

                {

                write-host $server "is already in MM" #display output#

                }

                else
                {
                write-host "$s Server not found" #display output#
                }
            }

        }

        Else
        {

        $SkipNo = 100*($i-1)

        $SkipBatch = Get-Content -Path $file | select -Skip $SkipNo

        Write-Host "Will be creating the batch $i " -ForegroundColor Yellow

        $Batch = $SkipBatch | select -First 100

        $Batch >> "C:\Softwares\$i.txt"

        Write-Host "Going to next batch for object $SkipNo"

            foreach ($s in $Batch)

            {
            $server = $instance | where {($_.DisplayName -like $s) -or ($_.DisplayName -contains $s+".contoso.local") } #selecting the windows computer object of each server#

                if($server.InMaintenanceMode -like "False") #Neglecting servers which are already in MM#

                {
                Start-SCOMMaintenanceMode -Reason PlannedApplicationMaintenance -Comment "Planned Maintenance" -Instance $server -EndTime $endtime

                write-host $server "has been kept in MM till $endtime" #display output#
                }

                Elseif($server.InMaintenanceMode -like "True")

                {

                write-host $server "is already in MM" #display output#

                }

                else
                {
                write-host "$s Server not found" #display output#
                }
            }
        }
    }
} 

