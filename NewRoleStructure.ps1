[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#Install-Module -Name AzureAD
Import-Module  AzureAD -UseWindowsPowerShell
$SiteURL = "https://yourdomain.sharepoint.com/sites/yoursite/"
$FolderSiteRelativeURL = "/Sites/yoursite/yourpathtodocuments/"
$GroupsToAvoid = "Group To Avoid", "Another One", "Another One"
$ListName = "Document Library"

#Connect to the Site
Connect-PnPOnline -URL $SiteURL -Interactive

#$Credential = Get-Credential
Connect-AzureAD -Credential $Credential

#Get the web & folder
$Web = Get-PnPWeb
$Folder = Get-PnPFolder -Url $FolderSiteRelativeURL
 
#Function to delete all Files and sub-folders from a Folder
Function ChangeRole-PnPFolder([Microsoft.SharePoint.Client.Folder]$Folder) {
    #Get the site relative path of the Folder
    If ($Web.ServerRelativeURL -eq "/") {
        $FolderSiteRelativeURL = $Folder.ServerRelativeUrl
    }
    Else {        
        $FolderSiteRelativeURL = $Folder.ServerRelativeUrl.Replace($Web.ServerRelativeURL, [string]::Empty)
    }
 
    #Delete all files in the Folder
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType File
    ForEach ($File in $Files) {
        #Change File
        Get-PnPPermissions $File.ListItemAllFields       
    }
 
    #Process all Sub-Folders
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder
    Foreach ($SubFolder in $SubFolders) {
        #Exclude "Forms" and Hidden folders
        If (($SubFolder.Name -ne "Forms") -and (-Not($SubFolder.Name.StartsWith("_")))) {
            #Call the function recursively
            ChangeRole-PnPFolder -Folder $SubFolder
 
            #Change the folder
            $ParentFolderURL = $FolderSiteRelativeURL.TrimStart("/")
            Get-PnPPermissions $SubFolder.ListItemAllFields
        }
    }
}

#Function to Get Permissions Applied on a particular Object such as: Web, List, Library, Folder or List Item
Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object) {
    Try {
        #Get permissions assigned to the Folder
        Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments

        #Check if Object has unique permissions
        $HasUniquePermissions = $Object.HasUniqueRoleAssignments
        
        if ($HasUniquePermissions) {
            #Loop through each permission assigned and extract details
            Get-PnPProperty -ClientObject $Object -Property File, Folder
            
            Foreach ($RoleAssignment in $Object.RoleAssignments) { 
                #Get the Permission Levels assigned and Member
                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
        
                #Get the Principal Type: User, SP Group, AD Group
                $PermissionType = $RoleAssignment.Member.PrincipalType
                $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
    
                #Remove Limited Access
                $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access" }) -join ","
                If ($PermissionLevels.Length -eq 0) { Continue }
    
                #Check if it is a sharepoint group, if not will be azure group
                If ($PermissionType -eq "SharePointGroup" -and $RoleAssignment.Member.LoginName -notin $GroupsToAvoid) {
                    Write-Host -f Red (" SharepointGroup {0}" -f $RoleAssignment.Member.LoginName)
                    #Get Group Members
                    $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName
                    
                    #Leave Change Groups
                    If ($GroupMembers.count -eq 0) { Continue }
                    Write-Host -f Green (" Processing URL {0}{1} - SharePoint Group: {2}" -f $Object.File.ServerRelativeUrl, $Object.Folder.ServerRelativeUrl, $RoleAssignment.Member.LoginName)

                    ForEach ($User in $GroupMembers) {
                        If ($User.Title -notin $GroupsToAvoid) {
                            If (!($RoleAssignment.Member.LoginName -like "*SharingLinks*" -and $PermissionLevels -eq "Read")) {
                                #Add the Data to Object
                                $Permissions = New-Object PSObject
                                $Permissions | Add-Member NoteProperty Name($User.Title)
                                $Permissions | Add-Member NoteProperty Type($PermissionType)
                                $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                                $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
                                $Permissions | Add-Member NoteProperty File($Object.File)
                                $Permissions | Add-Member NoteProperty Folder($Object.Folder)
                                $Permissions.Name = $Permissions.Name.Replace("'", "''")
                                SetSharepointRoleByAzureUserWithoutRemove($Permissions)
                            }                           
                        } 
                    }
                    
                    #Time to remove the old permission
                    if ($Object.File.ServerRelativeURL) {
                        Set-PnPListItemPermission -List $ListName -Identity $Object.File.ListItemAllFields -Group $RoleAssignment.Member.LoginName -RemoveRole $PermissionLevels -ErrorAction Stop
                    }
                    elseif ($Object.Folder) {
                        Set-PnPListItemPermission -List $ListName -Identity $Object.Folder.ListItemAllFields -Group $RoleAssignment.Member.LoginName -RemoveRole $PermissionLevels -ErrorAction Stop
                    }
                    
                }
                Elseif ($RoleAssignment.Member.Title -notin $GroupsToAvoid) {
                    If (!($RoleAssignment.Member.LoginName -like "*SharingLinks*" -and $PermissionLevels -eq "Read")) {
                        #Add the Data to Object
                        $Permissions = New-Object PSObject
                        $Permissions | Add-Member NoteProperty Name($RoleAssignment.Member.Title)
                        $Permissions | Add-Member NoteProperty Type($PermissionType)
                        $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
                        $Permissions | Add-Member NoteProperty File($Object.File)
                        $Permissions | Add-Member NoteProperty Folder($Object.Folder)
                        $Permissions.Name = $Permissions.Name.Replace("'", "''")

                        # Azure Group
                        If ((Get-AzureADGroup -Filter "DisplayName eq '$($Permissions.Name)'") -and $($Permissions.Name) -notlike "R *") {
                            Write-Host -f Green (" Processing URL {0}{1} | AzureAD Group: {2} | {3}" -f $Object.File.ServerRelativeUrl, $Object.Folder.ServerRelativeUrl, $Permissions.Name, $Permissions.Permissions)
                            SetSharepointRoleByAzureGroup($Permissions)

                        } # Azure User
                        elseif ((Get-AzureADUser -Filter "DisplayName eq '$($Permissions.Name)'") -and $($Permissions.Name) -notlike "*z_archive*") {
                            Write-Host -f Green (" Processing URL {0}{1} | AzureAD User: {2} | {3}" -f $Object.File.ServerRelativeUrl, $Object.Folder.ServerRelativeUrl, $Permissions.Name, $Permissions.Permissions)
                            SetSharepointRoleByAzureUser($Permissions)
                        }
                    }                           
                }
                    
            }
        }
        
    }
    Catch {
        write-host -f Red "Error:" $_.Exception.Message
    }
}

#
Function EnginePermissions($Object) {
    Write-Host -f Green (" URL {0}{1}" -f $Object.File.ServerRelativeUrl, $Object.Folder.ServerRelativeUrl)
    Write-Host -f Green (" '{0}' - '{1}'" -f $Object.Name, $Object.Permissions)

    Try {
        $Object.Name = $Object.Name.Replace("'", "''")

        # Azure Group
        If ((Get-AzureADGroup -Filter "DisplayName eq '$($Object.Name)'") -and $($Object.Name) -notlike "R *") {
            SetSharepointRoleByAzureGroup($Object)

        } # Azure User
        elseif ((Get-AzureADUser -Filter "DisplayName eq '$($Object.Name)'") -and $($Object.Name) -notlike "*z_archive*") {
            SetSharepointRoleByAzureUser($Object)
        }
    }
    Catch {
        write-host -f Red "Error!" $_.Exception.Message
    }
}

# Function to return the proper role, receive the object permissions, that include the name of the user.
Function Get-AzureProperRole() {
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$User
    )
    $aadGroupsOfUser = Get-AzureADUserMembership -ObjectId $($User.ObjectId) | where { $_.DisplayName -ne "All Users" }
    foreach ($group in $aadGroupsOfUser) {
        if ($group.DisplayName -like "R *") {
            Return $group
        }
    }
}

Function SetSharepointFilePermission() {
    Try {
        $AllFiles = Get-PnPFolderItem -FolderSiteRelativeUrl '/Shared Documents/Strategic Planning/Test' -ItemType File -ErrorAction Stop
    }
    Catch {
        Write-Host "Failed to list the files for '$($FolderName)'" -ForegroundColor Red
    }

    if ($AllFiles.count -ne 0) {
        Foreach ($File in $AllFiles) {
            try {
                if ($File.ServerRelativeUrl -eq "/sites/thebite/Shared Documents/Strategic Planning/Test/test.docx") {
                    Write-Host "Encontrado '$($File.Name)'" -ForegroundColor Red
                    #Set-PnPListItemPermission -List 'Document Library' -Identity $File.ListItemAllFields -User 'ray.deeks@birchandwaite.com.au' -AddRole 'Contribute' -ErrorAction Stop
                    Set-PnPListItemPermission -List 'Document Library' -Identity $File.ListItemAllFields -User 'c:0t.c|tenant|971febb8-1091-44d7-962b-c3e157430e94' -AddRole 'Contribute' -ErrorAction Stop
                }
            } 
            Catch {
                Write-Host "Folder $($FolderName): Failed to apply permissions to file $($File.Name). Error: $_.Exception.Message" -ForegroundColor Red
            }
        }
    }
    Else {
        Write-Host "'$($FolderName)' does not have any files" -ForegroundColor Yellow
    }
}

Function SetSharepointRoleByAzureUser() {
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$Object
    )
    if ((Get-AzureADUser -Filter "DisplayName eq '$($Object.Name)'") -and $($Object.Name) -notlike "*z_archive*") {
        $UserToAdd = Get-AzureADUser -Filter "DisplayName eq '$($Object.Name)' and UserType eq 'Member'"
        if ($UserToAdd.Name -notlike "*z_archive*") { 
            $Role = Get-AzureProperRole($UserToAdd)
            # Check if it is file, and foreach over each permission
            if ($Object.File.ServerRelativeURL -and $Role) {
                $PermissionList = $Object.Permissions.Split(",")
                foreach ($PermissionItem in $PermissionList) {
                    Set-PnPListItemPermission -List $ListName -Identity $Object.File.ListItemAllFields -User "c:0t.c|tenant|$($Role.ObjectId)" -AddRole $PermissionItem -ErrorAction Stop
                    Set-PnPListItemPermission -List $ListName -Identity $Object.File.ListItemAllFields -User "c:0t.c|tenant|$($UserToAdd.ObjectId)" -RemoveRole $PermissionItem -ErrorAction Stop
                }
                
            }
            elseif ($Object.Folder -and $Role) {
                $PermissionList = $Object.Permissions.Split(",")
                foreach ($PermissionItem in $PermissionList) {
                    Set-PnPListItemPermission -List $ListName -Identity $Object.Folder.ListItemAllFields -User "c:0t.c|tenant|$($Role.ObjectId)" -AddRole $PermissionItem -ErrorAction Stop
                    Set-PnPListItemPermission -List $ListName -Identity $Object.Folder.ListItemAllFields -User "c:0t.c|tenant|$($UserToAdd.ObjectId)" -RemoveRole $PermissionItem -ErrorAction Stop            
                }
            }
            else { Write-Host -f Red (" Rol not detected for the user {0}." -f $UserToAdd.DisplayName) }
        }
    }
}

function SetSharepointRoleByAzureGroup {    
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$Object
    )
    If ((Get-AzureADGroup -Filter "DisplayName eq '$($Object.Name)'") -and $($Object.Name) -notlike "R *") {
        # Get the members of the group, then find the proper role for each member, finally add the permission
        $GroupAzureAD = Get-AzureADGroup -Filter "DisplayName eq '$($Object.Name)'"
        $UserMembershipCollection = Get-AzureADGroupMember -ObjectId $GroupAzureAD.ObjectId -All $true | where { $_.DisplayName -notlike "*z_archive*" -and $_.ObjectType -ne "Group" }

        Foreach ($UserMember in $UserMembershipCollection) {
            $Role = Get-AzureProperRole($UserMember)
            if ($Object.File.ServerRelativeURL -and $Role) {
                #Write-Host -f Green ("      Role membership {0}:{1}:{2}" -f $Role.DisplayName, $UserMember.DisplayName, $GroupAzureAD.DisplayName)
                Set-PnPListItemPermission -List $ListName -Identity $Object.File.ListItemAllFields -User "c:0t.c|tenant|$($Role.ObjectId)" -AddRole $Object.Permissions -ErrorAction Stop
            }
            elseif ($Object.Folder -and $Role) {
                #Write-Host -f Green ("      Role membership {0}:{1}:{2}" -f $Role.DisplayName, $UserMember.DisplayName, $GroupAzureAD.DisplayName)
                Set-PnPListItemPermission -List $ListName -Identity $Object.Folder.ListItemAllFields -User "c:0t.c|tenant|$($Role.ObjectId)" -AddRole $Object.Permissions -ErrorAction Stop
            }               
                     
        }
        if ($Object.File.ServerRelativeURL) {
            Set-PnPListItemPermission -List $ListName -Identity $Object.File.ListItemAllFields -User "c:0t.c|tenant|$($GroupAzureAD.ObjectId)" -RemoveRole $Object.Permissions -ErrorAction Stop
        }
        elseif ($Object.Folder) {
            Set-PnPListItemPermission -List $ListName -Identity $Object.Folder.ListItemAllFields -User "c:0t.c|tenant|$($GroupAzureAD.ObjectId)" -RemoveRole $Object.Permissions -ErrorAction Stop
        }  

    }
}

Function SetSharepointRoleByAzureUserWithoutRemove() {
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$Object
    )
    if ((Get-AzureADUser -Filter "DisplayName eq '$($Object.Name)'") -and $($Object.Name) -notlike "*z_archive*") {
        $UserToAdd = Get-AzureADUser -Filter "DisplayName eq '$($Object.Name)' and UserType eq 'Member'"
        if ($UserToAdd.Name -notlike "*z_archive*") { 
            $Role = Get-AzureProperRole($UserToAdd)
            if ($Object.File.ServerRelativeURL -and $Role) {
                Set-PnPListItemPermission -List 'Document Library' -Identity $Object.File.ListItemAllFields -User "c:0t.c|tenant|$($Role.ObjectId)" -AddRole $($Object.Permissions) -ErrorAction Stop
            }
            elseif ($Object.Folder -and $Role) {
                Set-PnPListItemPermission -List 'Document Library' -Identity $Object.Folder.ListItemAllFields -User "c:0t.c|tenant|$($Role.ObjectId)" -AddRole $($Object.Permissions) -ErrorAction Stop               
            } 
        }
    }
}

Write-Host -f Green (" Starting the procedure...")

#Call the function to change role folder
ChangeRole-PnPFolder -Folder $Folder
