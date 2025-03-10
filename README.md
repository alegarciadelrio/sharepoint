# SharePoint Toolbox
🔭 A collection of useful PowerShell scripts to manage and maintain SharePoint sites.

<p>
  <img alt="PowerShell" src="https://img.shields.io/badge/PowerShell-black?style=flat-square&logoColor=white" />
  <img alt="SharePoint" src="https://img.shields.io/badge/Sharepoint-%23258AAF?style=flat-square&logo=sitepoint&logoColor=white" />
</p>

## Table of Contents
- [Overview](#overview)
- [Requirements](#requirements)
- [Repository Structure](#repository-structure)
- [Installation](#installation)
- [Usage](#usage)
  - [DeleteFolderSharepoint.ps1](#deletefolderSharepointps1)
  - [DeleteVersionsFileSharepoint.ps1](#deleteversionsfillesharePointps1)
  - [NewRoleStructure.ps1](#newrolestructureps1)
  - [RemovePerseverationHold.ps1](#removeperseverationholdps1)
  - [RemoveVersionHistory.ps1](#removeversionhistoryps1)
  - [RestoreRecycleBin.ps1](#restorerecyclebinps1)
- [Contributing](#contributing)
- [License](#license)

## Overview
This repository contains a set of PowerShell scripts designed to automate common SharePoint administration tasks. These scripts help streamline site management, version control, permission management, and content recovery operations.

## Requirements
- PowerShell 5.1 or higher
- SharePoint Online Management Shell
- Appropriate SharePoint permissions (Site Collection Administrator or equivalent)
- Microsoft 365 account with administrative access
- PnP PowerShell module (for some scripts)

## Repository Structure

This repository contains the following PowerShell scripts for SharePoint management:

| Script | Primary Function | Use Cases | Supported Features | Requirements |
|--------|-----------------|-----------|-------------------|--------------|
| **DeleteFolderSharepoint.ps1** | Delete SharePoint folders | • Content cleanup<br>• Site reorganization<br>• Removing obsolete content | • Recursive deletion<br>• Permission handling<br>• Detailed logging | • Site Collection Admin rights<br>• SharePoint Online Management Shell |
| **DeleteVersionsFileSharepoint.ps1** | Manage file version history | • Storage optimization<br>• Performance improvement<br>• Compliance management | • Selective version retention<br>• Major/minor version handling<br>• Space usage reporting | • Site Collection Admin rights<br>• SharePoint Online Management Shell |
| **NewRoleStructure.ps1** | Migrate to role-based permissions | • Security standardization<br>• Governance implementation<br>• Permission cleanup | • User-to-role mapping<br>• Group creation<br>• Permission migration<br>• Rollback support | • Site Collection Admin rights<br>• SharePoint Online Management Shell<br>• CSV mapping file |
| **RemovePerseverationHold.ps1** | Remove preservation holds | • Content lifecycle management<br>• Post-litigation cleanup<br>• Compliance operations | • Hold identification<br>• Selective removal<br>• Audit logging | • eDiscovery Admin rights<br>• Compliance Admin rights<br>• SharePoint Admin rights |
| **RemoveVersionHistory.ps1** | Bulk version history management | • Site-wide storage optimization<br>• Library maintenance<br>• Performance tuning | • Recursive processing<br>• Throttling controls<br>• Exclusion patterns<br>• Batch processing | • Site Collection Admin rights<br>• SharePoint Online Management Shell |
| **RestoreRecycleBin.ps1** | Restore deleted content | • Data recovery<br>• Accidental deletion recovery<br>• Content migration | • Folder hierarchy restoration<br>• Metadata preservation<br>• Conflict resolution<br>• Selective restoration | • Site Collection Admin rights<br>• SharePoint Online Management Shell |

All scripts follow consistent parameter naming conventions and error handling patterns to ensure reliability across different SharePoint environments. They can be used independently or as part of a larger workflow to address specific SharePoint administration challenges.

## Installation
1. Clone this repository or download the scripts to your local machine:
   ```
   git clone https://github.com/yourusername/sharepoint-toolbox.git
   ```
2. Ensure you have the required PowerShell modules installed:
   ```powershell
   Install-Module -Name Microsoft.Online.SharePoint.PowerShell
   Install-Module -Name PnP.PowerShell
   ```
3. Set the execution policy to allow running the scripts (if not already set):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage
Before running any script, you should connect to your SharePoint Online site:

```powershell
Connect-SPOService -Url https://yourtenant-admin.sharepoint.com
# or for PnP PowerShell
Connect-PnPOnline -Url https://yourtenant.sharepoint.com/sites/yoursite -Interactive
```

### DeleteFolderSharepoint.ps1
A powerful script for deleting folders from SharePoint document libraries. This script handles the complexities of SharePoint's folder structure and permissions, ensuring complete removal of the target folder and all its contents.

**Key Features:**
- Recursively deletes all subfolders and files
- Handles permission-inherited content
- Provides detailed logging of the deletion process
- Includes error handling for common SharePoint deletion issues
- Supports both modern and classic SharePoint sites

**Parameters:**
- `SiteUrl`: The URL of the SharePoint site containing the folder
- `FolderPath`: The relative path to the folder you want to delete
- `Confirm` (optional): When set to false, bypasses confirmation prompts
- `Verbose` (optional): Provides detailed operation logs

**Usage:**
```powershell
.\DeleteFolderSharepoint.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -FolderPath "/Shared Documents/FolderToDelete"
```

### DeleteVersionsFileSharepoint.ps1
This script manages version history for individual files in SharePoint, helping organizations optimize storage usage and improve site performance by removing unnecessary file versions while preserving essential revision history.

**Key Features:**
- Selectively removes old versions while keeping recent ones
- Preserves major versions while cleaning up minor versions if desired
- Supports all SharePoint file types (documents, images, etc.)
- Provides detailed reporting on space saved
- Respects file check-out status and handles locked files

**Parameters:**
- `SiteUrl`: The URL of the SharePoint site containing the file
- `FilePath`: The relative path to the file whose versions you want to manage
- `KeepVersions`: Number of most recent versions to retain
- `KeepMajorOnly` (optional): When true, removes all minor versions
- `Report` (optional): Generates a CSV report of actions taken

**Usage:**
```powershell
.\DeleteVersionsFileSharepoint.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -FilePath "/Shared Documents/YourFile.docx" -KeepVersions 5
```

### NewRoleStructure.ps1
A comprehensive permission management script that transforms individual user permissions into a more maintainable role-based security model. This script is essential for organizations scaling their SharePoint implementation or implementing governance policies.

**Key Features:**
- Analyzes existing permission structures
- Maps users to appropriate SharePoint groups based on their current permissions
- Creates new SharePoint groups if needed based on role definitions
- Migrates users to role-based groups while maintaining their effective permissions
- Removes direct permissions to improve security management
- Generates detailed reports of permission changes
- Supports rollback in case of migration issues

**Parameters:**
- `SiteUrl`: The URL of the SharePoint site to restructure
- `RoleMappingFile`: Path to a CSV file defining role mappings
- `CreateMissingGroups` (optional): Automatically creates groups that don't exist
- `ReportOnly` (optional): Analyzes and reports without making changes
- `IncludeSubsites` (optional): Applies changes to all subsites

**Usage:**
```powershell
.\NewRoleStructure.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -RoleMappingFile "roles.csv"
```

**Role Mapping File Format:**
```csv
UserEmail,CurrentPermission,TargetRole,TargetGroup
user@example.com,Contribute,Member,Site Members
admin@example.com,Full Control,Owner,Site Owners
```

### RemovePerseverationHold.ps1
This specialized script addresses compliance and retention policy challenges by removing preservation holds from SharePoint content. It's particularly useful for organizations that need to manage content lifecycle after legal or regulatory hold periods have expired.

**Key Features:**
- Identifies content with active preservation holds
- Safely removes holds without affecting content
- Works with both in-place holds and litigation holds
- Supports selective removal based on content age or metadata
- Provides audit logs for compliance reporting
- Handles both site-level and item-level holds

**Parameters:**
- `SiteUrl`: The URL of the SharePoint site containing held content
- `LibraryName`: The document library to process
- `HoldName` (optional): Specific hold identifier to remove
- `OlderThan` (optional): Only process items older than specified date
- `AuditLog` (optional): Path to save detailed audit information

**Usage:**
```powershell
.\RemovePerseverationHold.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -LibraryName "Documents"
```

**Note:** This script requires appropriate eDiscovery or Compliance administrator permissions in addition to SharePoint administrator rights.

### RemoveVersionHistory.ps1
A powerful bulk operation script that manages version history across entire folders or document libraries in SharePoint. This script is essential for storage optimization and performance improvement in SharePoint environments with extensive document versioning.

**Key Features:**
- Processes all files within a specified folder recursively
- Configurable version retention policy
- Intelligent handling of checked-out files
- Detailed logging of processing results
- Throttling controls to prevent performance impact
- Estimated storage savings calculation
- Support for excluding specific file types or patterns

**Parameters:**
- `SiteUrl`: The URL of the SharePoint site
- `FolderPath`: The relative path to the folder to process
- `KeepVersions`: Number of most recent versions to retain
- `Recursive` (optional): Process all subfolders when true
- `ExcludePattern` (optional): Skip files matching this pattern
- `BatchSize` (optional): Number of files to process in each batch
- `ReportOnly` (optional): Calculate potential savings without making changes

**Usage:**
```powershell
.\RemoveVersionHistory.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -FolderPath "/Shared Documents/YourFolder" -KeepVersions 3
```

### RestoreRecycleBin.ps1
A comprehensive data recovery script that simplifies the process of restoring deleted content from the SharePoint recycle bin. This script handles the complexities of restoring folder hierarchies with their contents while maintaining original metadata and permissions.

**Key Features:**
- Restores complete folder structures with all contained items
- Handles both first-stage and second-stage recycle bin items
- Preserves original metadata, version history, and permissions
- Resolves naming conflicts intelligently
- Provides detailed restoration reports
- Supports selective restoration based on deletion date or user
- Handles large restoration jobs with batch processing

**Parameters:**
- `SiteUrl`: The URL of the SharePoint site
- `FolderName`: Name of the folder to restore
- `RestorePoint` (optional): Restore items deleted before this date
- `DeletedBy` (optional): Only restore items deleted by this user
- `IncludeSubfolders` (optional): Restore all subfolders when true
- `ConflictResolution` (optional): How to handle naming conflicts
- `DetailedLog` (optional): Path for detailed restoration log

**Usage:**
```powershell
.\RestoreRecycleBin.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -FolderName "FolderToRestore"
```

**Advanced Usage:**
```powershell
.\RestoreRecycleBin.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -FolderName "FolderToRestore" -DeletedBy "user@example.com" -RestorePoint "2023-01-01" -ConflictResolution "CreateUnique"
```

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License
This project is licensed under the MIT License - see the LICENSE file for details.
