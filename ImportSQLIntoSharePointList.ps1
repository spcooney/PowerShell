##############################################################################################################
#
# Moves data from a SQL table into a SharePoint list
#
##############################################################################################################

# Global Variables (PLEASE UPDATE BEFORE RUNNING)

$dbConnStr = "Data Source=DBInstanceName;Initial Catalog=DBName;Integrated Security=True"
$siteUrl = "https://sharepointapppoolname/sites/sitename"

##############################################################################################################

# Load SharePoint Snapin
if (!(Get-PSSnapin | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" }))
{
	Write-Host "Adding the SharePoint PowerShell Snapin"
	Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

$ts = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"
$logName = ".\ImportLog_$ts.txt"

Start-Transcript -Path $logName

[System.Reflection.Assembly]::LoadWithPartialName("System.Data")
$dbConn = New-Object "System.Data.SqlClient.SqlConnection"
$dbConn.ConnectionString = $dbConnStr

$sw = [Diagnostics.Stopwatch]::StartNew()

Write-Host "Trying to connect to the database with the connection:" $dbConn.ConnectionString
try
{
	$dbConn.Open()
}
catch
{
	Write-Host "An error occurred trying to connect to the database.  Please ensure the connection string is correct and try again."
	Write-Host $_.Exception.Message
	exit
}
Write-Host "Connected to the database..."

# Create the select statement
$sqlCmd = New-Object System.Data.SqlClient.SqlCommand
$sqlCmd.Connection = $dbConn
$query = "SELECT * FROM [dbo].[RptTable] ORDER BY TableID"
$sqlCmd.CommandText = $query

# Create an adapter
$adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
$data = New-Object System.Data.DataSet
$adp.Fill($data) | Out-Null

$web = Get-SPWeb $siteUrl
if ($web -eq $null)
{
	Write-Host "Cannot find the site at: $siteUrl"
	exit
}

$spList = $web.Lists["SharePoint list title"];
if ($spList -eq $null)
{
	Write-Host "Cannot find the SharePoint list"
	exit
}
Write-Host "Starting to import the database rows into the SharePoint list"
Write-Host "============================================================================================="
Write-Host "This script will take 3 to 4 minutes to process all 4000+ records, so please be patient" -ForegroundColor red
Write-Host "This script should log every 100th item" -ForegroundColor red
Write-Host "============================================================================================="
$updCount = 0
foreach ($row in $data.Tables.Rows)
{ 
	try
	{
		$spQuery = New-Object Microsoft.SharePoint.SPQuery
		$caml = "<Where><Contains><FieldRef Name='ID'/><Value Type='Integer'>" + $row.ID + "</Value></Contains></Where>"
		$spQuery.Query = $caml
		$srchResult = $spList.GetItems($spQuery)
		if ($srchResult -ne $null -and $srchResult.Count -ge 0)
		{
			Write-Host "The following ID: " $row.ID " already exists"
			Continue
		}
		$item = $spList.AddItem();
		$item["Title"] = $row.ID
		if ($row.ID -ne $null -and $row.ID -isnot [DBNull])
		{
			$item["ID"] = [Int32]::Parse($row.ID)
		}
		if ($row.SomeString -ne $null)
		{
			$item["SomeString"] = $row.SomeString
		}
		if ($row.SomeDate -ne $null -and $row.SomeDate -isnot [DBNull])
		{
			$item["SomeDate"] = [DateTime]::Parse($row.SomeDate)
		}
		if ($item["SomeDate"] -eq '1/1/1900')
		{
			$item["SomeDate"] = $null
		}
		$legLinkURL = "/sites/sitename/Lists/SomeList/Forms/AllItems.aspx?FilterField1=ID&FilterValue1=" + $row.ID
		$item["UrlLink"] = "$legLinkURL, View Item"
		$item.Update()
		$updCount++
		if (($updCount % 100) -eq 0)
		{
			Write-Host "Currently processing list item number:" $updCount
		}
	}
	catch
	{
		Write-Host "ID: " $row.ID
		Write-Host "Exception Messsage: " $_.Exception.Message
		Write-Host "Stack Trace: " $_.Exception.StackTrace
		Write-Host "Line Number: " $_.InvocationInfo.ScriptLineNumber
	}
}
$sw.Stop()
Write-Host "Number of rows inserted: " $updCount
Write-Host "Number of list items: " $spList.Items.Count
Write-Host "The script took: " $sw.Elapsed.Minutes " minutes"

Stop-Transcript

$web.Dispose()