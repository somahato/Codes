#SDK Timeout increase

Import-Module OperationsManager

$ManagementServers = Get-SCOMManagementServer | ? {$_.IsGateway -eq $False}

$MSs = $ManagementServers.Displayname

Foreach ($MS in $MSs)

{

    ########################################################################################################
    # Registry Update                                                           #
    ########################################################################################################
   
   Invoke-Command -ComputerName $MS -ScriptBlock {reg add "HKLM\SOFTWARE\Microsoft\System Center\2010\Common\DAL" /v "DALCommandTimeoutSeconds" /t REG_DWORD /d 18000 /f}

}
