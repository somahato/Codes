<# How to run 64 bit command in 32-bit SCORCH activity run-net-script.

https://learn.microsoft.com/en-us/system-center/orchestrator/standard-activities/run-net-script?view=sc-orch-2022#execute-64-bit-powershell-cmdlets-with-run-net-script-activity

./ A sample example we could see but that is very generic and doesn't work all the time.

#>


$userName = "lab\sourav"
$password = ConvertTo-SecureString "Password" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList ($userName, $password)
$Result = Invoke-Command -Authentication Credssp -Credential $cred -ScriptBlock {
$Server = [Environment]::Is64BitProcess
$Version = $PSVersionTable.PSVersion
RETURN $Server , $Version

}  -ComputerName localhost

$Server = $Result[0]
$Version = $Result[1]
