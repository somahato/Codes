<#######################################################################################################################################################
# Author: Sourav Mahato
#
# Created Date:08/29/2023
#
# Purpose: Script To remove APM from SCOM completely
# Script to clean up obsolete references from Microsoft.SystemCenter.SecureReferenceOverride MP

 .Description
  This script connects to OpsMgr management group via SDK and detects obsolete management pack references in each unsealed management pack.

 .Parameter ManagementServer
  Specify the OpsMgr management server which the script connects to.

 .Parameter BackupBeforeModify
  Backup the unsealed management pack before making changes.

 .Parameter BackupLocation
  Specify the management pack backup location. This is a required parameter when -BackupBeforeModify switch is specified.

# https://learn.microsoft.com/en-us/troubleshoot/system-center/scom/remove-corrupted-apm-components
#
# How to run: PS C:\Script> .\Remove-APM.ps1 -ManagementServer scomms12019 -BackupBeforeModify "C:\SCOM"
OR
PS C:\Script> .\Remove-APM.ps1 -SDK "MS01" -BackupBeforeModify -BackupLocation "C:\SCOM"

#Got help from here https://blog.tyang.org/2014/06/24/powershell-script-remove-obsolete-references-unsealed-opsmgr-management-packs/
#######################################################################################################################################################>


Param (
	[Parameter(Mandatory=$true)][Alias("SDK")][string]$ManagementServer,
	[Parameter(Mandatory=$true)][switch]$BackupBeforeModify,
	[Parameter(Mandatory=$false)][string]$BackupLocation
)

#Region FunctionLibs
function Load-SDK()
{
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager.Common") | Out-Null
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager") | Out-Null
}

Function Get-ObsoleteReferences ([XML]$MPXML,[System.Collections.ArrayList]$arrCommonMPs)
{
	$arrObSoleteRefs = New-Object System.Collections.ArrayList
	$Refs = $MPXML.ManagementPack.Manifest.References.Reference
	Foreach ($Ref in $Refs)
	{
		$RefMPID = $Ref.ID
		$Alias = $Ref.Alias
		$strRef = "$Alias`!"
		If (!($MPXML.Innerxml.contains($strRef)))
		{
			#The alias is obsolete
			#Ignore Common MPs
			If (!($arrCommonMPs.contains($RefMPID)))
			{
				#Referecing MP is not a common MP
				[Void]$arrObSoleteRefs.Add($Ref)
			}

		}
	}
	,$arrObSoleteRefs
}

Function Get-MpXmlString ($MP)
{
    $MPStringBuilder = New-Object System.Text.StringBuilder
    $MPXmlWriter = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackXmlWriter([System.Xml.XmlWriter]::Create($MPStringBuilder))
    [Void]$MPXmlWriter.WriteManagementPack($MP)
    $MPXML = [XML]$MPStringBuilder.ToString()
    $MPXML
}

#EndRegion

Import-Module OperationsManager

New-SCOMManagementGroupConnection -ComputerName $ManagementServer

Write-Host "Removing Management Pack Operations Manager APM Reports Library" -ForegroundColor Yellow

Get-SCOMManagementPack -Name 'Microsoft.SystemCenter.DataWarehouse.ApmReports.Library' | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM Reports Library" -ForegroundColor Green

Write-Host "Removing Management Pack Operations Manager APM WCF Library" -ForegroundColor Yellow

Get-SCOMManagementPack -Name 'Microsoft.SystemCenter.Apm.Wcf' | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM WCF Library" -ForegroundColor Green

Write-Host "Removing Management Pack Operations Manager APM Web" -ForegroundColor Yellow

Get-SCOMManagementPack -Name 'Microsoft.SystemCenter.Apm.Web' | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM Web" -ForegroundColor Green

Write-Host "Removing Management Pack Operations Manager APM Windows Services" -ForegroundColor Yellow

Get-SCOMManagementPack -Name 'Microsoft.SystemCenter.Apm.NTServices' | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM Windows Services" -ForegroundColor Green

Write-Host "Removing Management Pack Operations Manager APM Infrastructure Monitoring" -ForegroundColor Yellow

Get-SCOMManagementPack -Name 'Microsoft.SystemCenter.Apm.Infrastructure.Monitoring' | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM Infrastructure Monitoring" -ForegroundColor Green

Write-Host "Removing Management Pack Operations Manager APM Library" -ForegroundColor Yellow

Get-SCOMManagementPack -Name 'Microsoft.SystemCenter.Apm.Library' | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM Library" -ForegroundColor Green

Write-Host "Removing Run as Profile association for APM" -ForegroundColor Yellow

$DWActionAccountProfile = Get-SCOMRunAsProfile -DisplayName "Data Warehouse Account"
$APMClass = Get-SCOMClass -DisplayName "Operations Manager APM Data Transfer Service"
$DWActionAccount = Get-SCOMrunAsAccount -Name "Data Warehouse Action Account"
If($APMClass){
Set-SCOMRunAsProfile -Action "Remove" -Profile $DWActionAccountProfile -Account $DWActionAccount -Class $APMClass}

Write-Host "Removed Run as Profile association for APM" -ForegroundColor Green

Write-Host "Modifying the XML Microsoft.SystemCenter.SecureReferenceOverride and removing Reference for Microsoft.SystemCenter.Apm.Infrastructure" -ForegroundColor Yellow

$arrCommonMPs = @("Microsoft.SystemCenter.Library",
"Microsoft.Windows.Library",
"System.Health.Library",
"System.Library",
"Microsoft.SystemCenter.DataWarehouse.Internal",
"Microsoft.SystemCenter.Notifications.Library",
"Microsoft.SystemCenter.DataWarehouse.Library",
"Microsoft.SystemCenter.OperationsManager.Library",
"System.ApplicationLog.Library",
"Microsoft.SystemCenter.Advisor.Internal",
"Microsoft.IntelligencePacks.Types",
"Microsoft.SystemCenter.Visualization.Configuration.Library",
"Microsoft.SystemCenter.Image.Library",
"Microsoft.SystemCenter.Visualization.ServiceLevelComponents",
"Microsoft.SystemCenter.NetworkDevice.Library",
"Microsoft.SystemCenter.InstanceGroup.Library",
"Microsoft.Windows.Client.Library")

#Connect to SCOM management group
Load-SDK
$MGConnSetting = New-Object Microsoft.EnterpriseManagement.ManagementGroupConnectionSettings($ManagementServer)
$MG = New-Object Microsoft.EnterpriseManagement.ManagementGroup($MGConnSetting)

#Create ManagementPackXMLWriter object if backup is required
If ($BackupBeforeModify)
{
	If (Test-Path $BackupLocation)
	{
		$date = Get-Date
		$BackupSubDir = "$($date.day)-$($date.month)-$($date.year) $($date.hour)`.$($date.minute)`.$($date.second)"
		$BackupDir = Join-Path $BackupLocation $BackupSubDir
	} else {
		Write-Error "Invalid Backup Location specified."
		Exit 2
	}
}

#Get all unsealed MPs
Write-Host "Getting Unsealed management packs..."
$strMPquery = "Name = 'Microsoft.SystemCenter.SecureReferenceOverride'"
$mpCriteria = New-Object  Microsoft.EnterpriseManagement.Configuration.ManagementPackCriteria($strMPquery)
$arrMPs = $MG.GetManagementPacks($mpCriteria)
Write-Host "Total number of unsealed management packs: $($arrMPs.count)" -ForegroundColor Yellow
Write-Host ""
$iTotalUpdated = 0
Foreach ($MP in $arrMPs)
{
	Write-Host "Checking MP: '$($MP.Name)'..." -ForegroundColor Green
	#Firstly, get the XML
	#$MPXML = Get-MPXML $MP.Name
	$MPXML = Get-MpXmlString $MP
	#Then get obsolete references (if there are any)
	$arrRefToDelete = Get-ObsoleteReferences $MPXML $arrCommonMPs
	
	If ($arrRefToDelete.count -gt 0)
	{
		
        Write-Host " - Number of obsolete references found: $($arrRefToDelete.count)" -ForegroundColor Yellow
		
		#Pre-Update MP verify
		Write-Host " - Verifying MP before updating it." -ForegroundColor Green
		Try
		{
			$bPreUpdateVerified = $MP.Verify()
		}Catch {
			$bPreUpdateVerified = $false
			Write-Host "   - MP verify failed. No changes will be made to this management pack. The MP Verify Error:" -ForegroundColor Red
			Write-Host $Error[0] -ForegroundColor Red
			Write-Host ""
		}
		
		If ($BackupBeforeModify -and $bPreUpdateVerified -ne $false)
		{
			If (!$WhatIf)
			{
				#Create Backup Dir if it's not present
                if (!(Test-path $BackupDir))
                {
                    New-Item -type directory -Path $BackupDir | Out-Null
                }

                #Create mpwriter if it's not created
                If (!$mpwriter)
                {
                    $mpWriter = new-object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackXmlWriter($BackupDir)
                }

                Write-Host " - Backing up $($MP.Name) to $BackupDir before modifying it." -ForegroundColor Yellow
				$mpWriter.WriteManagementPack($MP) | Out-Null
			} else {
				Write-Host " - $($MP.Name) would have been backed up to $BackupDir before modifying it." -ForegroundColor Red
			}
		}
		Foreach ($item in $arrRefToDelete)
		{	
			If (!$WhatIf -and $bPreUpdateVerified -ne $false)
			{
				Write-Host "  - Deleting reference '$($item.Alias)' `($($item.ID)`)" -ForegroundColor Yellow
				$MP.References.remove($item.Alias) | out-Null

			} else {
				if ($bPreUpdateVerified -ne $false)
				{
					Write-Host "  - The reference '$($item.Alias)' `($($item.ID)`) would have been deleted." -ForegroundColor Red
				} else {
					Write-Host "  - Pleae try manually remove the reference '$($item.Alias)' `($($item.ID)`)." -ForegroundColor Yellow
				}
			}
		}
				Try
				{
					$bPostUpdateVerified = $MP.Verify()
				}Catch {
					$bPostUpdateVerified = $false
				}

				#accept changes if MP is verified. otherwise reject changes
				If ($bPostUpdateVerified -eq $false)
				{
					Write-Host " - MP Verify failed. Reject changes." -ForegroundColor Red
					$MP.RejectChanges()
				} 
                else 
                {
					Write-Host " - MP Verified. Accepting changes." -ForegroundColor Yellow
					$MP.AcceptChanges()

				}		
	}
Else
{
Write-Host "- We didn't see any obsolete references: $($arrRefToDelete.count)" -ForegroundColor Green
}
	Write-Host ""
}

Write-Host "Done" -ForegroundColor Green
#endregion

Write-Host "Removing Management Pack Operations Manager APM Infrastructure" -ForegroundColor Yellow

Get-SCOMManagementPack -Name "Microsoft.SystemCenter.Apm.Infrastructure" | Remove-SCOMManagementPack

Write-Host "Removed Management Pack Operations Manager APM Infrastructure" -ForegroundColor Green
