##################################################################
#Author: Sourav Mahato
#Created Date:16/05/2024
#Purpose: Script to assign machines to DCRs (Data Collection Rules)
#How to run: You just need to provide the information like VMs, Subscription IDs and DCR details in the Script body.
##################################################################

# Install the AZ module
#Install-Module -Name Az -AllowClobber -Scope CurrentUser

#Provide the Azure Subscription ID
$SubscriptionId = "SubscriptionId"

#Connect to Azure Tenant
Connect-AzAccount -SubscriptionId $SubscriptionId

$ErrorActionPreference = 'SilentlyContinue'

$getRG = Get-AzResourceGroup #| Where-Object -FilterScript {$_.ResourceType -eq 'Microsoft.HybridCompute/machines'} | Select-Object ResourceGroupName -Unique
$ResourceGroupNames = $getRG.ResourceGroupName

#Provide the VM Names
$arcVirtualMachines = @('VM1','Test2','devtest_windows_event_dcr')

Foreach($ResourceGroupName in $ResourceGroupNames)

{
    #Provide the DCR Names
    $DataCollectionRuleNames = @('Test1','Test2','devtest_windows_event_dcr')
    Foreach($DataCollectionRuleName in $DataCollectionRuleNames)
    {
        $dcrNameList = (Get-AzDataCollectionRule -Name $DataCollectionRuleName -ResourceGroupName $ResourceGroupName).Id 
        If($dcrNameList)
        {    
            Foreach($arcVirtualMachine in $arcVirtualMachines)
            {
            
            $arcVirtualMachine = Get-AzResource -ResourceGroupName $ResourceGroupName -ResourceType Microsoft.Compute/virtualMachines -Name $arcVirtualMachine -ErrorAction SilentlyContinue
            if($arcVirtualMachine)        
            {
                $VM = $arcVirtualMachine.Name
                $VMId = $arcVirtualMachine.ResourceId
                $associationNames = (Get-AzDataCollectionRuleAssociation -DataCollectionRuleName $DataCollectionRuleName -ResourceGroupName $ResourceGroupName).Name
                If($associationNames)
                {
                    Foreach($associationName in $associationNames)                
                    {
                        $dataCollectionRuleId = (Get-AzDataCollectionRule -Name $DataCollectionRuleName -ResourceGroupName $ResourceGroupName).Id
                        New-AzDataCollectionRuleAssociation -AssociationName $associationName -ResourceUri $VMId -DataCollectionRuleId $dataCollectionRuleId -verbose
                    }
                }
                Else
                {
                    Write-Host "Didn't find any Association for the DCR $DataCollectionRuleName" -ForegroundColor Red
                }
            }
            Else
            {
                Write-Host "Didn't find any Virtual Machine on Resource Group $ResourceGroupName for the ResourceType Microsoft.Compute/virtualMachines" -ForegroundColor Red
            }
          }
        
       }
        Else
        {
            Write-Host "Didn't find the DataCollectionRuleName $DataCollectionRuleName on Resource Group $ResourceGroupName" -ForegroundColor Red
        }
    }
}
