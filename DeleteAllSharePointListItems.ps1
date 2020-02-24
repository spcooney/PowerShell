##############################################################################################################
#
# Deletes all of the SharePoint list data in the legacy list
#
##############################################################################################################

# Global Variables (PLEASE UPDATE BEFORE RUNNING)

$siteUrl = "https://sharepointapppoolname/sites/sitename"

##############################################################################################################

# Load SharePoint Snapin
if (!(Get-PSSnapin | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" }))
{
	Write-Host "Adding the SharePoint PowerShell Snapin"
	Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

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
$delCount = 0
foreach ($row in $spList.items)
{ 
	try
	{
		Write-Host "Deleting List Item ID:" $row["ID"]
		$spList.GetItemByID($row.id).Delete()
		$delCount++
	}
	catch
	{
		Write-Host $_.Exception.Message
	}
}
Write-Host "Number of rows deleted: $delCount"
$web.Dispose()