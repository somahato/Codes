#PowerShell Grid Widget dashboard

# https://learn.microsoft.com/en-us/system-center/scom/manage-create-web-dashboard-posh?view=sc-om-1807

#Script 1:

$class = Get-SCOMClass -Name Microsoft.Windows.Computer 
$Computers = @()
Get-SCOMClassInstance -Class $class | ForEach-Object{$computers += $_.DisplayName} 
$i=1 
$computers = Get-SCOMClassInstance -Class $class
Foreach ($computer in $computers)
{ 
    $perfRules = $computer.GetMonitoringPerformanceData() #| ? {$_.CounterName -like 'Semaphore Acquires' -Or $_.CounterName -like 'Average Semaphore Hold Time'}
    Foreach ($perfRule in $perfRules)
    {
    $dat1 = ($perfRule.GetValues($from, $now) | % { $_.SampleValue } | Measure-Object -Average).Average
    $Value = [Math]::Round($dat1,2) 
    $results=$ScriptContext.CreateFromObject($computer,"Id=Id,State=HealthState,DisplayName=DisplayName",$null)
    $results["Id"] = $computer.ToString()
    $dataObject["CustomColumn"]=$i
    $results["Path"] = $computer.FullName
    $results["DisplayName"] = $computer.DisplayName
    $results["CounterName"] = $perfRule.CounterName
    $results["Instancename"] = $perfRule.Instancename
    $results["ObjectName"] = $perfRule.ObjectName
    $results["Average"] = $Value
    $ScriptContext.ReturnCollection.Add($results) 
    $i++ 
    }
}

#Script 2:

$class = Get-SCOMGroup -Displayname "Test2"  
$Computers = @()
Get-SCOMClassInstance -Class $class | ForEach-Object{$computers += $_.DisplayName} 
$i=1 
$computers = Get-SCOMClassInstance -Group $class
Foreach ($computer in $computers)
{ 
    $perfRules = $computer.GetMonitoringPerformanceData() #| ? {$_.CounterName -like 'Semaphore Acquires' -Or $_.CounterName -like 'Average Semaphore Hold Time'}
    Foreach ($perfRule in $perfRules)
    {
    $dat1 = ($perfRule.GetValues($from, $now) | % { $_.SampleValue } | Measure-Object -Average).Average
    $Value = [Math]::Round($dat1,2) 
    $results=$ScriptContext.CreateFromObject($computer,"Id=Id,State=HealthState,DisplayName=DisplayName",$null)
    $results["Id"] = $computer.ToString()
    $results["Path"] = $computer.FullName
    $results["DisplayName"] = $computer.DisplayName
    $results["CounterName"] = $perfRule.CounterName
    $results["Instancename"] = $perfRule.Instancename
    $results["ObjectName"] = $perfRule.ObjectName
    $results["Average"] = $Value
    $ScriptContext.ReturnCollection.Add($results) 
    $i++ 
    }
}

#Populating Dashboard from a CSV file:

$File = "C:\Temp\SCOMMPList.csv"
$computers = @()
import-csv $File | ForEach-Object{$computers += $_.Name}
$Computers
$i = 1
foreach ($computer in $computers)  
{  
  $Datas = import-csv $File
  Foreach ($Data in $Datas)
  { 
  If ($Data.Name -like $computer)
  {
$dataObject = $ScriptContext.CreateInstance("xsd://foo!bar/baz")
$dataObject["Id"] = $computer.ToString()
$dataObject["CustomColumn"]=$i
$dataObject["Name"] = $computer
$dataObject["Version"] = $Data.Version
$dataObject["DisplayName"] = $Data.DisplayName
$dataObject["Sealed"] = $Data.Sealed
$ScriptContext.ReturnCollection.Add($dataObject) 
$i++
}
} 
}

# Populating Dashboard from a text file

$File = "C:\Test.log" 
$inputobject =Get-Content $File | ForEach{ "{0,5} {1}" -f $_.ReadCount,$_ }
        foreach ($line in $inputobject)
        {
            $dataObject = $ScriptContext.CreateInstance("xsd://foo!bar/baz")
            $dataObject["Id"] = $line.ToString()
            $dataobject["Log"]= $line
            $ScriptContext.ReturnCollection.Add($dataObject)
        }

# SCOM Agent Server Health state

$class = Get-SCOMClass -Name Microsoft.Windows.Computer 
$Computers = @()
Get-SCOMClassInstance -Class $class | ForEach-Object{$computers += $_.DisplayName} 
$i=1 
foreach ($computer in $computers) 
{
$Datas = Get-SCOMClassInstance -Class $class
Foreach ($Data in $Datas)
{ 
If ($Data.DisplayName -like $computer)
{
$dataObject=$ScriptContext.CreateInstance("xsd://foo!bar/baz")
$dataObject["Id"] = $computer.ToString()
$dataObject["CustomColumn"]=$i 
$dataObject["Name"] = $computer
$dataObject["Path"] = $Data.FullName
$dataObject["IsAvailable"] = $Data.IsAvailable
$dataObject["DisplayName"] = $Data.DisplayName
$dataObject["InMaintenanceMode"] = $Data.InMaintenanceMode 
$ScriptContext.ReturnCollection.Add($dataObject) 
$i++ 
}
}
}
