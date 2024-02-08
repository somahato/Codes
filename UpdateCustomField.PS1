param (
 
  [string]$AlertID
 
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 
# Import the Operations Manager Module
 
Import-Module OperationsManager
 
"Starting script"
 
$Alert = Get-SCOMAlert -Id $AlertID
 
$statement = $alert.Description

# Define a regular expression pattern to match the server name
$pattern = '\b(\w+\.)\b'

# Use the regex match to extract the server name
$match = [regex]::Match($statement, $pattern)

# Check if a match is found
if ($match.Success) {
    # Access the captured group (server name) using Groups[1]
    $serverName = $match.Groups[1].Value
    $serverName = $serverName.Split('.')
    Write-Host "Server Name: $serverName"
} else {
    Write-Host "Server name not found in the statement."
}

$Alert | Set-SCOMAlert -CustomField9 "$serverName"
