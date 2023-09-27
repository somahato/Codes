$Group = Get-SCOMGroup -DisplayName "All Windows COmputers" | Get-SCOMClassInstance

Function Get-SCOMContainedObjects {

   Param ([Microsoft.EnterpriseManagement.Monitoring.MonitoringObject[]]$ClassInstances)

   Foreach ($ClassInstance in $ClassInstances){

      Try{

          $ContainedHash.add($ClassInstance.ID,$ClassInstance)
       }

      Catch{}

      If ($Contained = $ClassInstance.GetRelatedMonitoringObjects()){

          Get-SCOMContainedObjects -ClassInstances $Contained
       }

   }
}

Function Get-SCOMContainedMonitoredObjects {

   Param ([Microsoft.EnterpriseManagement.Monitoring.MonitoringObject[]]$ClassInstances)

   $ContainedHash =@{}

   Get-SCOMContainedObjects -ClassInstances $ClassInstances 

   Return $ContainedHash.Values | Select-Object DisplayName,Fullname | Out-GridView

} 

Get-SCOMContainedMonitoredObjects -ClassInstances $Group
