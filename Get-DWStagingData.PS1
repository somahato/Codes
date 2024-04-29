##################################################################
#Author: Sourav Mahato
#Created Date:04/05/2024
#Purpose: Script get the staging area data from SCOM DataWarehouse Database
#How to run: You just need to provide the information as asked.
##################################################################

$SqlAuthLogin = Read-host ("Enter User Name")
$SqlAuthPw = Read-host ("Enter Password")
$SqlServer = Read-host ("SQL Server Instance Name")
$Database = Read-host ("Enter SCOM Data Warehouse Database Name")
$strSQLCommand = @("SELECT * FROM ManagedEntityStage","SELECT * FROM ManagedEntityStage",
"SELECT * FROM HealthServiceOutageStage",
"SELECT * FROM MaintenanceModeStage",
"SELECT * FROM RelationshipStage"
"SELECT * FROM TypedManagedEntityStage",
"SELECT * FROM Alert.AlertStage",
"SELECT *FROM Alert.AlertStage2Process",
"SELECT *FROM Event.EventStage",
"SELECT * FROM Perf.PerformanceStage",
"SELECT * FROM State.StateStage" )

Write-Host "***********************************************************************" -ForegroundColor DarkMagenta
Foreach ($SQLCommand in $strSQLCommand)
{
	# Build SQL Server connect string.
	$strSQLConnect = "Server=$SqlServer; Database=$Database; User Id=$SqlAuthLogin; Password=$SqlAuthPw;"
	# Connect to server and recieve dataset.
	$objSQLConnection = New-Object System.Data.SQLClient.SQLConnection
	$objSQLConnection.ConnectionString =  $strSQLConnect
	$objSQLCmd = New-Object System.Data.SQLClient.SQLCommand
	$objSQLCmd.CommandTimeout = $intSQLTimeout
	$objSQLCmd.CommandText = $SQLCommand
	$objSQLCmd.Connection = $objSQLConnection
	$objSQLAdapter = New-Object System.Data.SQLClient.SQLDataAdapter
	$objSQLAdapter.SelectCommand = $objSQLCmd
	$objDataSet = New-Object System.Data.DataSet
	$strRowCount = $objSQLAdapter.Fill($objDataSet)
       
  Write-Host "Count for Query '$SQLCommand' is $strRowCount" -ForegroundColor Yellow
  Write-Host "***********************************************************************" -ForegroundColor DarkMagenta
}

if ($?) {
		#pull out table
		$objTable = $objDataSet.Tables[0]
	}
	
# Close the SQL connection.
$objSQLConnection.Close()
	
# Return the dataset to caller.
return $objTable
