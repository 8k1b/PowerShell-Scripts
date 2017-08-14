## =====================================================================
## Title       : Run-Sql-Script
## Description : Runs all .sql scripts in sub-folder against specified 
##				 SQL Server instance and exports results to .csv for 
##				 each result set.  
## Author      : Jesse Reich
## Date        : 8/13/2017
## Input       : None
## Output      : .csv file for each result set of each .sql script in .\Scripts
## Usage	   : PS> .\Run-Sql-Script
## Notes 	   : User must configure global variales $ServerName  and $DatabaseName
##			   : User also must create a sub folder to hold .sql scripts
##			   : *** GO statements not allowed in scripts *** 
## Tag		   : SQL Server, WMI, Port, Configuration
## =====================================================================

$LogFile = ".\OutputLog.txt"
$ScriptFolder = ".\Scripts"

$ServerName = "LAPTOP-1"
$DatabaseName = "master"

# All message and error logging
Function Log-Message {
	Param([string]$Message, [bool]$IsError = $False)

	# Add the message to the log file
	Add-content $Logfile -value $Message
	
	# Output the message; if its an error, quit
	If($IsError) {
		Write-Error $Message -ErrorAction Stop
	}
	Else {
		Write-Output $Message
	}		
}

# Run Sql script code and export result sets to CSV files
Function Execute-Sql($Connection, $Sql, $FilePath){

	Log-Message -Message "Processing $FilePath..."

	$Command = New-Object System.Data.SqlClient.SqlCommand($Sql, $Connection)
	
	$DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $Command
	
	$Dataset = New-Object System.Data.Dataset
	
	$DataAdapter.Fill($Dataset) | Out-Null 
	
	# For each table in the dataset, export the results to a csv file
	For ($i=0; $i -lt $DataSet.Tables.Count; $i++)
	{
		$OutputFile = "{0}.result{1}.csv" -f $FilePath, $i
	
		$DataSet.Tables[$i] | Export-Csv $OutputFile -NoTypeInformation
	} 
}

# Program starts here 

# If the log file exists, clear its contents
If (Test-path $LogFile) {
	Clear-Content $LogFile 
}

# If the Scripts folder does not exist, log an error
If (!(Test-path $ScriptFolder)) {
	Log-Message -Message "The script folder is missing" -IsError $True
}

# Check if any scripts in the script folder
$ScriptSearch = "{0}\*.sql" -f $ScriptFolder

If (!(Test-Path $ScriptSearch)) {
	Log-Message -Message "There are no .sql files in the script folder" -IsError $True
}

# Main processing block
Try 
{
	# Connect to the database 
	$SqlConnection=New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$Database;Integrated Security=True"

	$SqlConnection.Open()

	# Loop through script folder for all .sql files
	Get-ChildItem $ScriptFolder -Filter *.sql | Foreach-Object {

		# Use Get-Contents to get T-SQL in file, must use .FullName for full path
		$Sql = Get-Content $_.FullName
	
		Execute-Sql -Connection $SqlConnection -Sql $Sql -FilePath $_.FullName
	}	

}
Catch 
{
	$SqlConnection.Close()

	# Throw exception to logger
	Log-Message -Message $_.Exception.Message -IsError $True
}
Finally
{
	# Always close database connection
	$SqlConnection.Close()
}

Log-Message -Message "Execution finished"


