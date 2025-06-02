# SharePoint File Size Analyzer Script - Enhanced with Owner and Access Information
# Identifies the largest files and analyzes total version history sizes

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    [Parameter(Mandatory=$false)]
    [string]$LibraryName = "Documents",
    [Parameter(Mandatory=$true)]
    [string]$LocalOutputPath,
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    [int]$RetryAttempts = 0,
    [int]$TopFilesCount = 5000,        # Number of largest files to analyze
    [int]$MinFileSizeMB = 1,           # Only analyze files larger than this
    [int]$SearchBatchSize = 500,       # Search batch size
    [string[]]$FileExtensions = @(),   # Optional: Filter by file extensions (e.g., @("pptx","docx","xlsx"))
    [switch]$IncludeAllFileTypes       # Include all file types if no extensions specified
)

# Ensure required modules
$requiredModules = @("PnP.PowerShell", "ImportExcel")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing required module: $module" -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
            Write-Host "Successfully installed $module" -ForegroundColor Green
        }
        catch {
            Write-Error "$module module installation failed. Please install manually with: Install-Module -Name $module -Scope CurrentUser"
            exit
        }
    }
}

# Global variables for tracking
$script:FilesAnalyzed = 0
$script:TotalCurrentSize = 0
$script:TotalVersionSize = 0
$script:FilesWithVersions = 0
$script:StartTime = Get-Date
$script:AnalysisResults = @()

# Create output directory
if (-not (Test-Path $LocalOutputPath)) {
    New-Item -ItemType Directory -Path $LocalOutputPath -Force | Out-Null
}

$logPath = Join-Path $LocalOutputPath "file_size_analysis_log.txt"
$csvOutputPath = Join-Path $LocalOutputPath "sharepoint_file_analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$excelOutputPath = Join-Path $LocalOutputPath "sharepoint_file_analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $logPath -Value $logMessage
}

function Get-LargestFilesUsingSearch {
    Write-Log "Searching for largest files using SharePoint Search..."
    
    try {
        $minSizeBytes = $MinFileSizeMB * 1024 * 1024
        
        # Build search query
        $searchQuery = "Size>=$minSizeBytes"
        
        # Add file extension filter if specified
        if ($FileExtensions.Count -gt 0) {
            $extensionFilter = ($FileExtensions | ForEach-Object { "FileExtension:$_" }) -join " OR "
            $searchQuery += " AND ($extensionFilter)"
        }
        
        Write-Log "Search Query: $searchQuery"
        Write-Log "Looking for files larger than $MinFileSizeMB MB..."
        Write-Log "Using paginated search with batch size: $SearchBatchSize"
        
        $allFiles = @()
        $startRow = 0
        $totalRetrieved = 0
        
        do {
            Write-Log "Searching batch starting at row $startRow..."
            
            # Execute search with pagination - Enhanced with more properties for owner/access info
            $searchResults = Submit-PnPSearchQuery -Query $searchQuery -StartRow $startRow -MaxResults $SearchBatchSize -SelectProperties "Title,Path,Size,LastModifiedTime,FileType,FileExtension,Author,Created,ModifiedBy,CreatedBy,ViewsLifeTime,ViewsRecent,LastViewedTime,CheckoutUser"
            
            if ($searchResults.ResultRows.Count -eq 0) {
                Write-Log "No more results found in this batch"
                break
            }
            
            Write-Log "Retrieved $($searchResults.ResultRows.Count) results in this batch"
            
            foreach ($result in $searchResults.ResultRows) {
                try {
                    # Skip folders and system files
                    if ($result.FileType -eq "aspx" -or $result.Path -like "*/_*" -or $result.Path -like "*/Forms/*") {
                        continue
                    }
                    
                    $fileInfo = [PSCustomObject]@{
                        Name = [System.IO.Path]::GetFileName($result.Path)
                        ServerRelativeUrl = $result.Path -replace "^https?://[^/]+", ""
                        SizeMB = [math]::Round([long]$result.Size / 1MB, 2)
                        SizeBytes = [long]$result.Size
                        TimeLastModified = [datetime]$result.LastModifiedTime
                        TimeCreated = [datetime]$result.Created
                        FileExtension = $result.FileExtension
                        Author = $result.Author
                        CreatedBy = $result.CreatedBy
                        ModifiedBy = $result.ModifiedBy
                        ViewsLifeTime = $result.ViewsLifeTime
                        ViewsRecent = $result.ViewsRecent
                        LastViewedTime = if ($result.LastViewedTime) { [datetime]$result.LastViewedTime } else { $null }
                        CheckoutUser = $result.CheckoutUser
                        FullPath = $result.Path
                    }
                    
                    $allFiles += $fileInfo
                    $totalRetrieved++
                }
                catch {
                    Write-Log "Error processing search result: $($_.Exception.Message)" "WARN"
                }
            }
            
            $startRow += $SearchBatchSize
            
            # Brief pause between batches to avoid throttling
            Start-Sleep -Seconds 1
            
        } while ($searchResults.ResultRows.Count -eq $SearchBatchSize)
        
        if ($allFiles.Count -eq 0) {
            Write-Log "No files found matching criteria" "WARN"
            return @()
        }
        
        # Remove duplicates and sort by size (largest first)
        $allFiles = $allFiles | Sort-Object ServerRelativeUrl -Unique | Sort-Object SizeBytes -Descending
        
        # Take top N files
        $topFiles = $allFiles | Select-Object -First $TopFilesCount
        
        Write-Log "Found $($allFiles.Count) total files, analyzing top $($topFiles.Count) largest files"
        Write-Log "Total size of files to analyze: $([math]::Round(($topFiles | Measure-Object SizeMB -Sum).Sum, 2)) MB"
        
        # Show top 10 largest files for context
        if ($topFiles.Count -gt 0) {
            Write-Log "Top 10 largest files:"
            $topFiles | Select-Object -First 10 | ForEach-Object {
                Write-Log "  - $($_.Name): $($_.SizeMB) MB ($($_.FileExtension))"
            }
        }
        
        return $topFiles
    }
    catch {
        Write-Log "Search failed: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Get-EnhancedFileInfo {
    param([object]$File)
    
    try {
        # Get detailed file information including owner and permissions
        $fileDetails = Get-PnPFile -Url $File.ServerRelativeUrl -AsListItem -ErrorAction SilentlyContinue
        
        $ownerInfo = @{
            FileOwner = "Unknown"
            FileOwnerEmail = "Unknown"
            LastAccessedBy = $File.ModifiedBy
            LastAccessedDate = $File.TimeLastModified
        }
        
        if ($fileDetails) {
            # Try to get owner information from various fields
            if ($fileDetails["Author"]) {
                $ownerInfo.FileOwner = $fileDetails["Author"].LookupValue
                if ($fileDetails["Author"].Email) {
                    $ownerInfo.FileOwnerEmail = $fileDetails["Author"].Email
                }
            }
            elseif ($fileDetails["Created_x0020_By"]) {
                $ownerInfo.FileOwner = $fileDetails["Created_x0020_By"].LookupValue
                if ($fileDetails["Created_x0020_By"].Email) {
                    $ownerInfo.FileOwnerEmail = $fileDetails["Created_x0020_By"].Email
                }
            }
            
            # Get last modified by information
            if ($fileDetails["Modified_x0020_By"]) {
                $ownerInfo.LastAccessedBy = $fileDetails["Modified_x0020_By"].LookupValue
            }
            elseif ($fileDetails["Editor"]) {
                $ownerInfo.LastAccessedBy = $fileDetails["Editor"].LookupValue
            }
            
            # Use Modified date as last accessed date
            if ($fileDetails["Modified"]) {
                $ownerInfo.LastAccessedDate = $fileDetails["Modified"]
            }
        }
        
        # If we still don't have owner info, use the original Author field
        if ($ownerInfo.FileOwner -eq "Unknown" -and $File.Author) {
            $ownerInfo.FileOwner = $File.Author
        }
        
        # If we still don't have owner info, use CreatedBy
        if ($ownerInfo.FileOwner -eq "Unknown" -and $File.CreatedBy) {
            $ownerInfo.FileOwner = $File.CreatedBy
        }
        
        Write-Log "Enhanced info for $($File.Name): Owner=$($ownerInfo.FileOwner), LastAccessed=$($ownerInfo.LastAccessedBy)"
        
        return $ownerInfo
        
    } catch {
        Write-Log "Could not retrieve enhanced file info for $($File.Name): $($_.Exception.Message)" "WARN"
        
        # Return fallback information
        return @{
            FileOwner = if ($File.Author) { $File.Author } elseif ($File.CreatedBy) { $File.CreatedBy } else { "Unknown" }
            FileOwnerEmail = "Unknown"
            LastAccessedBy = if ($File.ModifiedBy) { $File.ModifiedBy } else { "Unknown" }
            LastAccessedDate = $File.TimeLastModified
        }
    }
}

function Get-FileVersionsWithRetry {
    param(
        [string]$FileUrl,
        [string]$FileName,
        [int]$RetryCount = 0
    )
    
    try {
        $versions = Get-PnPFileVersion -Url $FileUrl -ErrorAction Stop
        Write-Log "Found $($versions.Count) versions for: $FileName"
        return $versions
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Failed to get versions for $FileName`: $errorMsg" "WARN"
        
        # Check for throttling
        if ($errorMsg -like "*throttle*" -or $errorMsg -like "*429*" -or $errorMsg -like "*rate limit*") {
            Write-Log "Throttling detected. Waiting 30 seconds..." "WARN"
            Start-Sleep -Seconds 30
        }
        
        if ($RetryCount -lt $RetryAttempts) {
            $waitTime = [Math]::Pow(2, $RetryCount) * 5
            Write-Log "Retrying in $waitTime seconds..." "INFO"
            Start-Sleep -Seconds $waitTime
            return Get-FileVersionsWithRetry -FileUrl $FileUrl -FileName $FileName -RetryCount ($RetryCount + 1)
        }
        
        Write-Log "Unable to retrieve versions for $FileName after $RetryAttempts attempts" "ERROR"
        return @()
    }
}

function Analyze-FileVersions {
    param([object]$File)
    
    try {
        Write-Log "Analyzing: $($File.Name) ($($File.SizeMB) MB)"
        
        # Get enhanced file information (owner, last accessed, etc.)
        $enhancedInfo = Get-EnhancedFileInfo -File $File
        
        # Get all versions for this file
        $versions = Get-FileVersionsWithRetry -FileUrl $File.ServerRelativeUrl -FileName $File.Name
        
        $currentSize = $File.SizeBytes
        $totalVersionSize = $currentSize  # Include current version
        $versionCount = 1  # Current version
        $oldestVersionDate = $File.TimeCreated
        $versionSizes = @()
        
        if ($versions.Count -gt 0) {
            $script:FilesWithVersions++
            $versionCount += $versions.Count
            
            foreach ($version in $versions) {
                $totalVersionSize += $version.Size
                $versionSizes += [math]::Round($version.Size / 1MB, 2)
                
                if ($version.Created -lt $oldestVersionDate) {
                    $oldestVersionDate = $version.Created
                }
            }
        }
        
        $totalVersionSizeMB = [math]::Round($totalVersionSize / 1MB, 2)
        $versionOverheadMB = [math]::Round(($totalVersionSize - $currentSize) / 1MB, 2)
        $versionOverheadPercent = if ($currentSize -gt 0) { [math]::Round((($totalVersionSize - $currentSize) / $currentSize) * 100, 1) } else { 0 }
        
        # Create analysis result with enhanced owner and access information
        $analysisResult = [PSCustomObject]@{
            FileName = $File.Name
            FileExtension = $File.FileExtension
            ServerRelativeUrl = $File.ServerRelativeUrl
            CurrentSizeMB = $File.SizeMB
            TotalVersionCount = $versionCount
            TotalSizeAllVersionsMB = $totalVersionSizeMB
            VersionOverheadMB = $versionOverheadMB
            VersionOverheadPercent = $versionOverheadPercent
            FileOwner = $enhancedInfo.FileOwner
            FileOwnerEmail = $enhancedInfo.FileOwnerEmail
            LastAccessedBy = $enhancedInfo.LastAccessedBy
            LastAccessedDate = $enhancedInfo.LastAccessedDate.ToString('yyyy-MM-dd HH:mm:ss')
            OriginalAuthor = $File.Author
            CreatedBy = $File.CreatedBy
            ModifiedBy = $File.ModifiedBy
            ViewsLifeTime = if ($File.ViewsLifeTime) { $File.ViewsLifeTime } else { 0 }
            ViewsRecent = if ($File.ViewsRecent) { $File.ViewsRecent } else { 0 }
            LastViewedTime = if ($File.LastViewedTime) { $File.LastViewedTime.ToString('yyyy-MM-dd HH:mm:ss') } else { "Never" }
            CheckoutUser = if ($File.CheckoutUser) { $File.CheckoutUser } else { "None" }
            Created = $File.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            LastModified = $File.TimeLastModified.ToString('yyyy-MM-dd HH:mm:ss')
            OldestVersion = $oldestVersionDate.ToString('yyyy-MM-dd HH:mm:ss')
            DaysOldestToNewest = [math]::Round(($File.TimeLastModified - $oldestVersionDate).TotalDays, 1)
            DaysSinceLastAccessed = [math]::Round(((Get-Date) - $enhancedInfo.LastAccessedDate).TotalDays, 1)
            AverageVersionSizeMB = if ($versionCount -gt 0) { [math]::Round($totalVersionSizeMB / $versionCount, 2) } else { 0 }
            LargestVersionMB = if ($versionSizes.Count -gt 0) { ($versionSizes | Measure-Object -Maximum).Maximum } else { $File.SizeMB }
            SmallestVersionMB = if ($versionSizes.Count -gt 0) { ($versionSizes.Count -gt 0) ? ($versionSizes | Measure-Object -Minimum).Minimum : $File.SizeMB } else { $File.SizeMB }
            FullPath = $File.FullPath
        }
        
        $script:AnalysisResults += $analysisResult
        $script:TotalCurrentSize += $currentSize
        $script:TotalVersionSize += $totalVersionSize
        
        Write-Log "✓ Analysis complete: $versionCount versions, $totalVersionSizeMB MB total ($versionOverheadMB MB overhead) - Owner: $($enhancedInfo.FileOwner)" "SUCCESS"
        
        return $analysisResult
    }
    catch {
        Write-Log "Error analyzing $($File.Name): $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Export-AnalysisResults {
    Write-Log "Exporting analysis results..."
    
    try {
        # Sort results by total version size (largest first)
        $sortedResults = $script:AnalysisResults | Sort-Object TotalSizeAllVersionsMB -Descending
        
        # Export to CSV
        $sortedResults | Export-Csv -Path $csvOutputPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report exported to: $csvOutputPath" "SUCCESS"
        
        # Export to Excel with formatting and summary
        $excelPackage = $sortedResults | Export-Excel -Path $excelOutputPath -WorksheetName "File Analysis" -AutoSize -FreezeTopRow -BoldTopRow -PassThru
        
        # Add summary worksheet
        $summaryData = @(
            [PSCustomObject]@{Metric = "Total Files Analyzed"; Value = $script:FilesAnalyzed}
            [PSCustomObject]@{Metric = "Files with Version History"; Value = $script:FilesWithVersions}
            [PSCustomObject]@{Metric = "Files without Versions"; Value = ($script:FilesAnalyzed - $script:FilesWithVersions)}
            [PSCustomObject]@{Metric = "Total Current Size (GB)"; Value = [math]::Round($script:TotalCurrentSize / 1GB, 2)}
            [PSCustomObject]@{Metric = "Total All Versions Size (GB)"; Value = [math]::Round($script:TotalVersionSize / 1GB, 2)}
            [PSCustomObject]@{Metric = "Total Version Overhead (GB)"; Value = [math]::Round(($script:TotalVersionSize - $script:TotalCurrentSize) / 1GB, 2)}
            [PSCustomObject]@{Metric = "Version Overhead Percentage"; Value = [math]::Round((($script:TotalVersionSize - $script:TotalCurrentSize) / $script:TotalCurrentSize) * 100, 1)}
        )
        
        $summaryData | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Summary" -AutoSize -BoldTopRow
        
        # Add top files by version overhead worksheet
        $topOverhead = $sortedResults | Where-Object { $_.VersionOverheadMB -gt 0 } | Sort-Object VersionOverheadMB -Descending | Select-Object -First 50
        if ($topOverhead.Count -gt 0) {
            $topOverhead | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Top Version Overhead" -AutoSize -BoldTopRow
        }
        
        # Add files by owner analysis
        $ownerAnalysis = $sortedResults | Group-Object FileOwner | ForEach-Object {
            $group = $_.Group
            [PSCustomObject]@{
                FileOwner = $_.Name
                FileCount = $_.Count
                TotalCurrentSizeMB = [math]::Round(($group | Measure-Object CurrentSizeMB -Sum).Sum, 2)
                TotalAllVersionsSizeMB = [math]::Round(($group | Measure-Object TotalSizeAllVersionsMB -Sum).Sum, 2)
                AverageVersionsPerFile = [math]::Round(($group | Measure-Object TotalVersionCount -Average).Average, 1)
                TotalVersionOverheadMB = [math]::Round(($group | Measure-Object VersionOverheadMB -Sum).Sum, 2)
                AverageDaysSinceLastAccess = [math]::Round(($group | Measure-Object DaysSinceLastAccessed -Average).Average, 1)
            }
        } | Sort-Object TotalAllVersionsSizeMB -Descending
        
        $ownerAnalysis | Export-Excel -ExcelPackage $excelPackage -WorksheetName "By Owner" -AutoSize -BoldTopRow
        
        # Add stale files analysis (files not accessed in 90+ days)
        $staleFiles = $sortedResults | Where-Object { $_.DaysSinceLastAccessed -gt 90 } | Sort-Object DaysSinceLastAccessed -Descending
        if ($staleFiles.Count -gt 0) {
            $staleFiles | Select-Object FileName, FileOwner, LastAccessedBy, LastAccessedDate, DaysSinceLastAccessed, CurrentSizeMB, TotalSizeAllVersionsMB, VersionOverheadMB | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Stale Files (90+ days)" -AutoSize -BoldTopRow
        }
        
        # Add file type analysis
        $fileTypeAnalysis = $sortedResults | Group-Object FileExtension | ForEach-Object {
            $group = $_.Group
            [PSCustomObject]@{
                FileExtension = $_.Name
                FileCount = $_.Count
                TotalCurrentSizeMB = [math]::Round(($group | Measure-Object CurrentSizeMB -Sum).Sum, 2)
                TotalAllVersionsSizeMB = [math]::Round(($group | Measure-Object TotalSizeAllVersionsMB -Sum).Sum, 2)
                AverageVersionsPerFile = [math]::Round(($group | Measure-Object TotalVersionCount -Average).Average, 1)
                TotalVersionOverheadMB = [math]::Round(($group | Measure-Object VersionOverheadMB -Sum).Sum, 2)
                UniqueOwners = ($group | Select-Object -ExpandProperty FileOwner | Sort-Object -Unique).Count
                AverageDaysSinceLastAccess = [math]::Round(($group | Measure-Object DaysSinceLastAccessed -Average).Average, 1)
            }
        } | Sort-Object TotalAllVersionsSizeMB -Descending
        
        $fileTypeAnalysis | Export-Excel -ExcelPackage $excelPackage -WorksheetName "By File Type" -AutoSize -BoldTopRow
        
        Close-ExcelPackage $excelPackage
        Write-Log "Excel report exported to: $excelOutputPath" "SUCCESS"
        
        return @{
            CSVPath = $csvOutputPath
            ExcelPath = $excelOutputPath
            TotalFiles = $script:FilesAnalyzed
            TotalCurrentSizeGB = [math]::Round($script:TotalCurrentSize / 1GB, 2)
            TotalVersionSizeGB = [math]::Round($script:TotalVersionSize / 1GB, 2)
        }
    }
    catch {
        Write-Log "Error during export: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Show-Progress {
    param([int]$Current, [int]$Total, [string]$CurrentFile)
    
    $elapsed = (Get-Date) - $script:StartTime
    $percentComplete = [math]::Round(($Current / $Total) * 100, 1)
    
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "PROGRESS: $Current/$Total files analyzed ($percentComplete%)" -ForegroundColor Yellow
    Write-Host "Current: $CurrentFile" -ForegroundColor White
    Write-Host "Files with Versions: $($script:FilesWithVersions)" -ForegroundColor Green
    Write-Host "Current Size Total: $([math]::Round($script:TotalCurrentSize / 1GB, 2)) GB" -ForegroundColor Cyan
    Write-Host "All Versions Total: $([math]::Round($script:TotalVersionSize / 1GB, 2)) GB" -ForegroundColor Magenta
    Write-Host "Version Overhead: $([math]::Round(($script:TotalVersionSize - $script:TotalCurrentSize) / 1GB, 2)) GB" -ForegroundColor Red
    Write-Host "Elapsed Time: $($elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
    if ($Current -gt 0) {
        $eta = $elapsed.TotalSeconds * ($Total - $Current) / $Current
        $etaSpan = [TimeSpan]::FromSeconds($eta)
        Write-Host "ETA: $($etaSpan.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
    }
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host ""
}

# Main execution
Write-Log "Starting SharePoint File Size Analysis with Owner and Access Information..."
Write-Log "Site URL: $SiteUrl"
Write-Log "Library: $LibraryName"
Write-Log "Output Path: $LocalOutputPath"
Write-Log "Top Files Count: $TopFilesCount"
Write-Log "Minimum File Size: $MinFileSizeMB MB"
Write-Log "File Extensions Filter: $(if ($FileExtensions.Count -gt 0) { $FileExtensions -join ', ' } else { 'All file types' })"

try {
    # Connect to SharePoint
    Write-Log "Connecting to SharePoint site..."
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    Write-Log "Successfully connected to SharePoint!"
    
    # Get largest files
    Write-Log "Starting file discovery..."
    $largestFiles = Get-LargestFilesUsingSearch
    
    if ($largestFiles.Count -eq 0) {
        Write-Log "No files found matching criteria" "WARN"
        exit
    }
    
    # Analyze each file
    $fileIndex = 0
    foreach ($file in $largestFiles) {
        $fileIndex++
        
        Show-Progress -Current $fileIndex -Total $largestFiles.Count -CurrentFile $file.Name
        
        try {
            $result = Analyze-FileVersions -File $file
            $script:FilesAnalyzed++
        }
        catch {
            Write-Log "Error analyzing $($file.Name): $($_.Exception.Message)" "ERROR"
        }
        
        # Brief pause to avoid throttling
        Start-Sleep -Seconds 2
    }
    
    # Export results
    Write-Log "Analysis complete. Exporting results..."
    $exportResult = Export-AnalysisResults
    
    # Final summary
    $totalElapsed = (Get-Date) - $script:StartTime
    $versionOverheadGB = [math]::Round(($script:TotalVersionSize - $script:TotalCurrentSize) / 1GB, 2)
    $overheadPercent = if ($script:TotalCurrentSize -gt 0) { [math]::Round((($script:TotalVersionSize - $script:TotalCurrentSize) / $script:TotalCurrentSize) * 100, 1) } else { 0 }
    
    Write-Log ""
    Write-Log "========== ANALYSIS COMPLETE =========="
    Write-Log "Files Analyzed: $($script:FilesAnalyzed)"
    Write-Log "Files with Version History: $($script:FilesWithVersions) ($([math]::Round(($script:FilesWithVersions / $script:FilesAnalyzed) * 100, 1))%)"
    Write-Log "Total Current File Size: $([math]::Round($script:TotalCurrentSize / 1GB, 2)) GB"
    Write-Log "Total All Versions Size: $([math]::Round($script:TotalVersionSize / 1GB, 2)) GB"
    Write-Log "Version History Overhead: $versionOverheadGB GB ($overheadPercent%)"
    Write-Log "Processing Time: $($totalElapsed.ToString('hh\:mm\:ss'))"
    Write-Log ""
    Write-Log "Reports Generated:"
    Write-Log "- CSV: $csvOutputPath"
    Write-Log "- Excel: $excelOutputPath"
    Write-Log "- Log: $logPath"
    Write-Log "============================================="
    
    # Show top 5 files by version overhead
    $topOverhead = $script:AnalysisResults | Where-Object { $_.VersionOverheadMB -gt 0 } | Sort-Object VersionOverheadMB -Descending | Select-Object -First 5
    if ($topOverhead.Count -gt 0) {
        Write-Log ""
        Write-Log "Top 5 Files by Version Overhead:"
        foreach ($file in $topOverhead) {
            Write-Log "- $($file.FileName): $($file.VersionOverheadMB) MB overhead ($($file.TotalVersionCount) versions) - Owner: $($file.FileOwner)"
        }
    }
    
    # Show stale files summary
    $staleFilesCount = ($script:AnalysisResults | Where-Object { $_.DaysSinceLastAccessed -gt 90 }).Count
    if ($staleFilesCount -gt 0) {
        Write-Log ""
        Write-Log "Found $staleFilesCount files not accessed in 90+ days (see 'Stale Files' worksheet)"
    }
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
}
finally {
    try {
        Disconnect-PnPOnline
        Write-Log "Disconnected from SharePoint"
    } catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" "ERROR"
    }
}