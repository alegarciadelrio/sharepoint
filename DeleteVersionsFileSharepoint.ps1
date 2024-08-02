#Parameters
$SiteURL = "https://yoursite.sharepoint.com/sites/yoursite/"
$FileURL= "/Sites/yoursite/Shared Documents/SALES/file.xlsm"

#Connect to PnP Online
Connect-PnPOnline -URL $SiteURL -Interactive
  
#Get all versions of the File
$Versions = Get-PnPFileVersion -Url $FileURL
 
#Delete all versions of the File
$Versions.DeleteAll()
Invoke-PnPQuery