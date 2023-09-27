$Username = 'lab\sqlsvc'
$password = Read-Host -AsSecureString

$NewCred = new-object System.Management.Automation.PsCredential $UserName,$Password

Get-SCOMRunAsAccount -Name "SQL Service Account" | Update-SCOMRunAsAccount -RunAsCredential $NewCred
