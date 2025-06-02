# FileAnalytics
A series of applets to analyse files

## OVERVIEW
This PowerShell script analyzes SharePoint document libraries to identify the largest files and calculate storage overhead from version history. It provides detailed information about file ownership, access patterns, and version management to help optimize SharePoint storage.

## KEY FEATURES
- Identifies largest files in SharePoint libraries using search
- Analyzes version history and calculates storage overhead
- Tracks file ownership and last access information
- Exports results to CSV and Excel with multiple analysis worksheets
- Provides progress tracking and detailed logging
- Handles throttling and retry logic
- Supports file type filtering

## REQUIREMENTS

### Software Prerequisites
- Windows PowerShell 5.1 or PowerShell 7+
- Internet connection
- SharePoint Online access

### PowerShell Modules
The script will automatically install these modules if missing:
- PnP.PowerShell (SharePoint connectivity)
- ImportExcel (Excel export functionality)

### SharePoint Permissions
The user running the script needs:
- Site Collection Administrator OR
- Site Owner permissions OR
- Read access to the target document library PLUS
- Search permissions on the SharePoint site

### Azure App Registration (Recommended)
For automated/scheduled execution, create an Azure App Registration with:
- SharePoint permissions: Sites.Read.All or Sites.FullControl.All
- Grant admin consent for the organization

## USAGE

### Basic Usage
```powershell
.\filesizeaudit.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -LocalOutputPath "C:\SharePointAnalysis" -ClientId "your-app-id" -TenantId "your-tenant-id"
```

### Advanced Usage with Options
```powershell
.\filesizeaudit.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite" -LibraryName "Documents" -LocalOutputPath "C:\SharePointAnalysis" -ClientId "your-app-id" -TenantId "your-tenant-id" -TopFilesCount 1000 -MinFileSizeMB 5 -FileExtensions @("pptx","docx","xlsx") -RetryAttempts 3
```

## PARAMETERS

### Required Parameters
- **SiteUrl**: Full URL to your SharePoint site
- **LocalOutputPath**: Local directory for output files
- **ClientId**: Azure App Registration Client ID
- **TenantId**: Your Office 365 Tenant ID

### Optional Parameters
- **LibraryName**: Document library name (default: "Documents")
- **TopFilesCount**: Number of largest files to analyze (default: 5000)
- **MinFileSizeMB**: Minimum file size in MB to analyze (default: 1)
- **SearchBatchSize**: Search results per batch (default: 500)
- **RetryAttempts**: Number of retry attempts for failed operations (default: 0)
- **FileExtensions**: Array of file extensions to filter (e.g., @("pptx","docx","xlsx"))
- **IncludeAllFileTypes**: Switch to include all file types when no extensions specified

## OUTPUT FILES

The script generates several output files in the specified LocalOutputPath:

### CSV Report
- **sharepoint_file_analysis_YYYYMMDD_HHMMSS.csv**: Complete analysis data

### Excel Report with Multiple Worksheets
- **File Analysis**: Complete file-by-file analysis
- **Summary**: Overall statistics and totals
- **Top Version Overhead**: Files with highest version storage overhead
- **By Owner**: Analysis grouped by file owner
- **Stale Files (90+ days)**: Files not accessed in 90+ days
- **By File Type**: Analysis grouped by file extension

### Log File
- **file_size_analysis_log.txt**: Detailed execution log

## ANALYSIS DATA FIELDS

Each analyzed file includes:

### File Information
- FileName, FileExtension, ServerRelativeUrl, FullPath
- CurrentSizeMB, TotalVersionCount, TotalSizeAllVersionsMB
- VersionOverheadMB, VersionOverheadPercent

### Ownership and Access
- FileOwner, FileOwnerEmail, OriginalAuthor, CreatedBy, ModifiedBy
- LastAccessedBy, LastAccessedDate, DaysSinceLastAccessed
- ViewsLifeTime, ViewsRecent, LastViewedTime, CheckoutUser

### Dates and Versions
- Created, LastModified, OldestVersion
- DaysOldestToNewest, AverageVersionSizeMB
- LargestVersionMB, SmallestVersionMB

## PERMISSIONS SETUP

### For Interactive Use (Recommended for Testing)
1. Run the script with your admin account
2. Use Interactive authentication (script will prompt for login)
3. Ensure you have appropriate SharePoint permissions

### For Automated Use (App Registration)
1. Go to Azure Portal > App Registrations
2. Create new registration:
   - Name: "SharePoint File Analyzer"
   - Supported account types: Single tenant
   - Redirect URI: Not required
3. Note the Application (client) ID and Directory (tenant) ID
4. Go to API Permissions:
   - Add Microsoft Graph: Sites.Read.All (Application)
   - Add SharePoint: Sites.FullControl.All or Sites.Read.All (Application)
   - Grant admin consent
5. Go to Certificates & secrets:
   - Create client secret (note the value)
6. Modify connection in script if using client secret instead of interactive

### Minimum SharePoint Permissions
- Read access to target document library
- Search permissions on the site
- Access to file version history

## CUSTOMIZATION POINTS

### Search Query Modification
Located in `Get-LargestFilesUsingSearch` function:
```powershell
$searchQuery = "Size>=$minSizeBytes"
```
Modify to add additional search criteria:
- Content type filters: `ContentType:"Document"`
- Path filters: `Path:"https://yourtenant.sharepoint.com/sites/yoursite/library/*"`
- Date filters: `LastModifiedTime>=2023-01-01`

### File Type Filtering
Modify the `$FileExtensions` parameter or add logic in the search loop:
```powershell
if ($result.FileType -eq "aspx" -or $result.Path -like "*/_*" -or $result.Path -like "*/Forms/*") {
    continue
}
```

### Progress Display Frequency
Modify the `Show-Progress` function call frequency:
```powershell
# Show progress every 10 files instead of every file
if ($fileIndex % 10 -eq 0) {
    Show-Progress -Current $fileIndex -Total $largestFiles.Count -CurrentFile $file.Name
}
```

### Analysis Criteria
Modify version analysis logic in `Analyze-FileVersions`:
- Change stale file threshold (currently 90 days)
- Add custom file size thresholds
- Include/exclude specific file paths

### Output Customization
Add custom worksheets in `Export-AnalysisResults`:
```powershell
# Add custom analysis worksheet
$customAnalysis = $sortedResults | Where-Object { /* custom criteria */ }
$customAnalysis | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Custom Analysis" -AutoSize -BoldTopRow
```

### Throttling and Performance
Adjust delays and batch sizes:
```powershell
$SearchBatchSize = 500        # Reduce if encountering throttling
Start-Sleep -Seconds 2        # Increase delay between file analyses
Start-Sleep -Seconds 30       # Throttling recovery delay
```

## TROUBLESHOOTING

### Common Issues

**"Access Denied" Errors**
- Verify SharePoint permissions
- Check if files are checked out or have unique permissions
- Ensure search permissions are granted

**"Throttling Detected"**
- Increase delays between operations
- Reduce batch sizes
- Run during off-peak hours

**"Module Not Found" Errors**
- Run PowerShell as Administrator
- Manually install modules: `Install-Module -Name PnP.PowerShell -Scope CurrentUser`

**"Authentication Failed"**
- Verify ClientId and TenantId are correct
- Check App Registration permissions
- Ensure admin consent is granted

### Performance Optimization

**For Large Libraries (10,000+ files)**
- Increase MinFileSizeMB to focus on larger files
- Reduce TopFilesCount for initial analysis
- Run analysis in smaller batches by file type

**For Slow Connections**
- Increase SearchBatchSize to reduce API calls
- Add longer delays between operations
- Consider running during off-peak hours

## SECURITY CONSIDERATIONS

### Sensitive Information
- The script accesses file metadata and version information
- Owner and access information is collected
- Ensure output files are stored securely
- Consider data retention policies for generated reports

### Credentials
- Never hard-code credentials in the script
- Use Azure App Registration for automated scenarios
- Secure storage of Client Secrets if using app-only authentication
- Regular rotation of authentication credentials

## SUPPORT AND MAINTENANCE

### Regular Updates
- Monitor for PnP.PowerShell module updates
- Review SharePoint API changes that might affect search
- Update Azure App Registration permissions as needed

### Monitoring
- Review log files for patterns of failures
- Monitor for changes in file access patterns
- Track storage optimization progress over time

For technical support or script modifications, consult your SharePoint administrator or PowerShell developer.

## VERSION HISTORY
- v1.0: Initial release with enhanced owner/access tracking
- Includes comprehensive version analysis and multi-worksheet Excel output
