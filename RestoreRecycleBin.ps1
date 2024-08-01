#Parameter
$SiteURL = "https://yoursite.sharepoint.com/sites/yoursite/"
$DirPath = "Sites/yoursite/Shared Documents/folder"
  
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
 
#Get the Web
$Web = Get-PnPWeb
  
#Get All Items deleted from a specific path or library - sort by most recently deleted
$DeletedItems = Get-PnPRecycleBinItem -RowLimit 100 | Where { $_.DirName -like "$DirPath*"} | Sort-Object -Property DeletedDate -Descending
 
#Restore all deleted items from the given path to its original location
ForEach($Item in $DeletedItems)
{
    #Get the Original location of the deleted file
    $OriginalLocation = "/"+$Item.DirName+"/"+$Item.LeafName
    $OriginalLocation
    If($Item.ItemType -eq "File")
    {
        $OriginalItem = Get-PnPFile -Url $OriginalLocation -AsListItem -ErrorAction SilentlyContinue
    }
    Else #Folder
    {
        $OriginalItem = Get-PnPFolder -Url $OriginalLocation -ErrorAction SilentlyContinue
    }
    #Check if the item exists in the original location
    If($OriginalItem -eq $null)
    { 
        #Restore the item
        $Item | Restore-PnpRecycleBinItem -Force
        Write-Host "Item '$($Item.LeafName)' restored Successfully!" -f Green
    }
    Else
    {
        Write-Host "There is another file with the same name.. Skipping $($Item.LeafName)" -f Yellow
    }
}