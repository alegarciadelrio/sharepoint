$SiteURL = "https://yoursite.sharepoint.com/sites/yoursite/"
$ListName = "Preservation Hold Library"

#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
 
#Delete all files from the library
Get-PnPList -Identity $ListName | Get-PnPListItem -PageSize 100 -ScriptBlock {
    Param($items) Invoke-PnPQuery } | ForEach-Object { $_.Recycle() | Out-Null
}