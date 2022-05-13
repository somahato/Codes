##################################################################
#Author: Sourav Mahato
#Created Date:01/27/2014
#Purpose: Script for Agent health fix
##################################################################
###### Function for SQL Query
#######################################################################

Function Get-SQLTable($strSQLServer, $strSQLDatabase, $strSQLCommand, $intSQLTimeout = 3000)
{	
	trap [System.Exception]
	{
		DisplayUpdate "Exception trapped, $($_.Exception.Message)"
		DisplayUpdate "SQL Command Failed.  Sql Server [$strSQLServer], Sql Database [$strSQLDatabase], Sql Command [$strSQLCommand]."
		continue;
	}
	
	#build SQL Server connect string
	$strSQLConnect = "Server=$strSQLServer;Database=$strSQLDatabase;Integrated Security=True;Connection Timeout=$intSQLTimeout" 
	
	#connect to server and recieve dataset
	$objSQLConnection = New-Object System.Data.SQLClient.SQLConnection
	$objSQLConnection.ConnectionString =  $strSQLConnect
	$objSQLCmd = New-Object System.Data.SQLClient.SQLCommand
	$objSQLCmd.CommandTimeout = $intSQLTimeout
	$objSQLCmd.CommandText = $strSQLCommand
	$objSQLCmd.Connection = $objSQLConnection
	$objSQLAdapter = New-Object System.Data.SQLClient.SQLDataAdapter
	$objSQLAdapter.SelectCommand = $objSQLCmd
	$objDataSet = New-Object System.Data.DataSet
	$strRowCount = $objSQLAdapter.Fill($objDataSet)
	
	If ($?)
	{
		#pull out table
		$objTable = $objDataSet.tables[0]
	}
	
	#close the SQL connection
	$objSQLConnection.Close()
	
	#return array of values to caller	
	return $objTable
}


#######################################################################
    ##     FUNCTION TO CHECK HEARTBEAT
#######################################################################
function Get-IsAvailable ($agent, $sql, $db)
{
	trap [System.Exception]
	{
		DisplayUpdate "Exception trapped, $($_.Exception.Message)"
		continue;
	}
	
	#connect to the SCOM zone SQL agent
	$cmd = "select fullName, a.IsAvailable, a.LastModified
	from BaseManagedEntity (nolock) as b
	join Availability (nolock) as a
	on b.BaseManagedEntityId = a.BaseManagedEntityId
	where b.FullName like 'Microsoft.SystemCenter.HealthService:$agent.%'"

	$healthService = Get-SqlTable $sql $db $cmd
	
	if ($healthService.IsAvailable -ne 0)
	{
		return [bool] $healthService.IsAvailable
	}
	else
	{
		return $false
	}
}

#######################################################################
##### FUNCTION FOR TO GET MANAGEMENT Server NAME FROM AGENT agent
#######################################################################

function Get-ManagementServer ($agent)
{
	$omsdkbranch = "LocalMachine"
    $omsdkregistry=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Localmachine',$agent)
	
    $omsdksubbranch = "SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Agent Management Groups"
    $omsdkRegistrykey = $omsdkregistry.OpenSubKey($omsdksubbranch)
    $Zone = $omsdkRegistrykey.GetSubKeyNames()
    Foreach ($zon in $zone)
    {
    
        if ($zon -match 'OM12_')
        {
            $zonename = $zon
        }
      
    }

    #HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Agent Management Groups\Zonename\Parent Health Services\0
	$hklm = 2147483650
	$regPath = "SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Agent Management Groups\$zonename\Parent Health Services\0"
	$regValue = "AuthenticationName"
	$regprov = [wmiclass]"\\$agent\root\default:stdRegProv"
	return ($regprov.GetStringValue($hklm,$regPath,$regValue)).svalue
}
#######################################################################
### FUNCTION FOR TO GET SQL agent
#######################################################################

function Get-ScomSqlagent ($ms)
{
	#HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup\DatabaseagentName
	$hklm = 2147483650
	$regPath = "SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup"
	$regValue = "DatabaseServerName"
	
	$regprov = [wmiclass]"\\$ms\root\default:stdRegProv"
	return ($regprov.GetStringValue($hklm,$regPath,$regValue)).svalue
}

#######################################################################
### FUNCTION FOR TO GET SQL DBNAME
#######################################################################

function Get-ScomDbName ($ms)
{
	#HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\setup\DatabaseName
	$hklm = 2147483650
	$regPath = "SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup"
	$regValue = "DatabaseName"
	
	$regprov = [wmiclass]"\\$ms\root\default:stdRegProv"
	return ($regprov.GetStringValue($hklm,$regPath,$regValue)).svalue
}

#######################################################################
    ##     Function to get Domain
#######################################################################

function getagentdomain($agent)  
{
$branch = "LocalMachine"
	$dnsdataregistry=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($branch,$agent)
	$dnsdatasubbranch = "SYSTEM\CurrentControlSet\services\Tcpip\Parameters"
	$dnsdataRegistrykey = $dnsdataregistry.OpenSubKey($dnsdatasubbranch)
	If ($dnsdataRegistrykey -ne $null) 
    {
		$agentdomain = $dnsdataRegistrykey.GetValue("Domain")
	}
	else 
    {
		$agentdomain = $null
	}
		return $agentdomain	
}

#######################################################################
    ##     Function to get zone info from SCAM
#######################################################################
function Check-ScamForZone ($zone)
{
	trap [System.Exception]
	{
		DisplayUpdate "Exception trapped, $($_.Exception.Message)"
		continue;
	}
	
	#SCAM information
	$scamServer = "SCAM"
	$scamDb = "SCAM"
	
	#build command
	$cmd = "select top 1 ManagementGroupName
			from ManagementGroups
			where ManagementGroupName = '$zone'"

	$scamInfo = Get-SqlTable $scamServer $scamDb $cmd
	
	if ($scamInfo.Count)
	{
		return $scamInfo[0].ManagementGroupName
	}
	elseif ($scamInfo.ManagementGroupName)
	{
		return $scamInfo.ManagementGroupName
	}
	else
	{
		return $false
	}
}

#################################################################################################
###Code start from here
#################################################################################################

$cur = Get-location

    if (Test-path $cur\log.txt)

        {
            Remove-Item -Path "$cur\log.txt" -Force
        }

$srvs = Get-content $cur\Agents.txt

        if (!$srvs)

            {
                $srvs_temp = Read-host ("Enter server names")
                
                $srvs = $srvs_temp.Split(',')

            }


foreach ($computer in $srvs)

{
    $agent=$computer.trim()
    
    if ($agent -match "STGOM")
    
    {

    "agent $agent is belong to SCOM MS or RMS agent">>$cur\Log.txt
    
    Write-Host "$agent is belong to SCOM MS or RMS agent" -ForegroundColor Red

    "****************************************************************************************************" >>$cur\Log.txt
    }
    
    else

    {
        $ms = Get-ManagementServer $agent

            if ( $ms -notmatch "STGOM" )

                {
                    "agent $agent is not point to SCOM MS or RMS agent, need to Install the Agent">>$cur\Log.txt

                    Write-Host "$agent is not point to SCOM MS or RMS agent, need to Install the Agent" -ForegroundColor Red
                       
                    "****************************************************************************************************" >>$cur\Log.txt
                }
        
                else

                    {
                        $sql = Get-ScomSqlagent $ms

                        if ($sql -notmatch "STGOM12SQ")

                        {
                            "agent $agent is not point to right prod sql agent">>$cur\Log.txt

                            Write-Host "$agent is not point to right prod sql agent" -ForegroundColor Red
                               
                            "****************************************************************************************************" >>$cur\Log.txt
                        }

                        else

                            {
                                $db = Get-ScomDbName $ms
    
                                if ($db -notmatch "OM12")

                                {
                                    "agent $agent is not point to right DB">>$cur\Log.txt

                                    Write-Host "$agent is not point to right DB" -ForegroundColor Red
                                       
                                    "****************************************************************************************************" >>$cur\Log.txt
                                }
               
                                else
                                  {
                                    $healthService = Get-IsAvailable $agent $sql $db
	
	                                if ($healthService -ne 0)
	                                {
                                        "agent $agent is already healthy" >>$cur\Log.txt

                                        Write-Host "$agent is already healthy" -ForegroundColor Green
                                           
                                        "****************************************************************************************************" >>$cur\Log.txt 
                                    }

                                    else

                                        {
                                            $servicename = "Healthservice"

                                            $status = "Running"

                                            $ServicetoStart= Get-Wmiobject -Class win32_service -computer $agent -filter "name = 'healthservice'"
               
                                            if ($ServicetoStart.state -ne $status)

                                            {
                                                Write-Host "Health Service is stopped on $agent, Started the Service" -ForegroundColor Yellow

                                                (Get-Wmiobject -Class win32_service -computer $agent -filter "name = 'healthservice'").InvokeMethod("StartService",$null)

                                                Start-Sleep -Seconds 30

                                                (Get-Wmiobject -Class win32_service -computer $agent -filter "name = 'healthservice'") >> $cur\Output.txt
               
                                            }
               
                                            else

                                                {   

                                                    Write-Host "Health Service is Running on $agent but $agent is Unhealthy on Console, Flush the Service" -ForegroundColor Yellow
                                                    
                                                    (Get-Wmiobject -Class win32_service -computer $agent -filter "name = 'healthservice'").InvokeMethod("StopService",$null)

                                                    Start-Sleep -Seconds 30

                                                    Rename-Item -path "\\$agent\c$\Program Files\System Center Operations Manager\Agent\Health Service State" -newname "Health Service State_Old"

                                                    (Get-Wmiobject -Class win32_service -computer $agent -filter "name = 'healthservice'").InvokeMethod("StartService",$null)
               
                                                    Start-Sleep -Seconds 30

                                                    (Get-Wmiobject -Class win32_service -computer $agent -filter "name = 'healthservice'") >> $cur\Output.txt

               #Once the Health Service State folder got created, its recommanded to remove old file, as it consume space and when u run next time it will fail to rename as same name folder exists
               
                                                    if (Test-Path "\\$agent\c$\Program Files\System Center Operations Manager\Agent\Health Service State")
                                                                                        
                                                       {
                                                            Remove-Item "\\$agent\c$\Program Files\System Center Operations Manager\Agent\Health Service State_Old" -Recurse
                                                       }

                                                }
        
                                        }
                                   }
                            }
                    }
    }
}
            
          

