
#Variable
$SiteURL= "//required site url//"
 
#Connect to PnP Online
Connect-PnPOnline $SiteURL -Credentials (Get-Credential)
 
#Get List Item count from all Lists from the Web
Get-PnPList | Select Title, ItemCount
