
#Variable
$SiteURL= "//required site url//"
 
#Connect to PnP Online
Connect-PnPOnline $SiteURL -Credentials (Get-Credential)
 
#Get List Item count from all Lists from the Web
Get-PnPList | Select Title, ItemCount

#in case of MFA issues use this
$SiteURL= "//required site url//"
Connect-PnPOnline $SiteURL -UseWebLogin
Get-PnPList | Select Title, ItemCount
