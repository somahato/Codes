$Agent = get-itemproperty -path "HKLM:\SOFTWARE\Microsoft\System Center Operations Manager\12\Setup\Agent" -ErrorAction SilentlyContinue ;$Agent.'RTM_UR Version'
$AgentVerion = $Agent.'RTM_UR Version'

If ($AgentVerion -eq '10.22.10056.0'){"SCOM 2022 RTM"}
Elseif ($AgentVerion -eq '10.22.10110.0'){"SCOM 2022 UR1"}
Elseif ($AgentVerion -eq '10.22.10208.0'){"SCOM 2022 UR2"}
Elseif ($AgentVerion -eq '10.22.10215.0'){"SCOM 2022 UR2 Hotfix"}
Elseif ($AgentVerion -eq '10.19.10014.0	'){"SCOM 2019 RTM"}
Elseif ($AgentVerion -eq '10.19.10140.0'){"SCOM 2019 UR1"}
Elseif ($AgentVerion -eq '10.19.10153.0'){"SCOM 2019 UR2"}
Elseif ($AgentVerion -eq '10.19.10177.0'){"SCOM 2019 UR3"}
Elseif ($AgentVerion -eq '10.19.10185.0'){"SCOM 2019 UR3 Hotfix"}
Elseif ($AgentVerion -eq '10.19.10200.0'){"SCOM 2019 UR4"}
Elseif ($AgentVerion -eq '10.19.10211.0'){"SCOM 2019 UR5"}
Elseif ($AgentVerion -eq '10.19.10253.0	'){"SCOM 2019 UR6"}

if(!$AgentVerion)
{

$Agent = get-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup" -ErrorAction SilentlyContinue ;$Agent.'CurrentVersion'
$AgentVerion = $Agent.'CurrentVersion'
If ($AgentVerion -eq '10.22.10056.0'){"SCOM 2022 RTM"}
Elseif ($AgentVerion -gt '10.19.10176.0' -and $AgentVerion -lt '10.19.99999.9'){"Supported Azure Agent"}
Elseif ($AgentVerion -gt '10.20.18053.0' -and $AgentVerion -lt '10.20.99999.9'){"Supported Azure Agent"}
else{"Agent is not installed"}
}
