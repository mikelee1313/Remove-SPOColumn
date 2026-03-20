
<#
    .SYNOPSIS
        Removes a specified column from all lists and libraries across multiple SharePoint Online sites.

    .DESCRIPTION
        This script connects to multiple SharePoint Online sites and removes a specified column (field)
        from every list and library where it exists. The script includes comprehensive throttling
        protection, retry logic, and detailed logging capabilities.
        
        The script processes each list and library in each site:
        - Checks whether the target column exists in each list/library
        - Removes the column if found, skipping lists where it is not present
        - Handles both site columns and list-level fields
        
        Throttling protection is implemented with configurable delays between operations and exponential
        backoff retry logic for HTTP 429/503 responses.

    .PARAMETER Verbose
        When specified, displays additional diagnostic information during processing.
        Useful for debugging and tracking column discovery across lists.

    .EXAMPLE
        .\Remove-SPOColumn.ps1
        
        Runs the script with default settings, removing the configured column from all sites in SPOSiteList.txt

    .EXAMPLE
        .\Remove-SPOColumn.ps1 -Verbose
        
        Runs the script with verbose output, showing additional details for each list processed

    .NOTES
        File Name      : Remove-SPOColumn.ps1
        Author         : Mike Lee
        Date           : 3/20/26
        Prerequisite   : PnP.PowerShell module, Entra App Registration with certificate authentication
        
        Required Permissions:
        - Sites.ReadWrite.All (or Sites.FullControl.All)
        - User.Read.All
        
        Configuration:
        Before running, update the USER CONFIGURATION section with:
        - App ID from your Entra App Registration
        - Tenant ID
        - Certificate thumbprint
        - Tenant admin URL (e.g. https://contoso-admin.sharepoint.com) — required for full tenant scan
        - Path to input file containing site URLs (one per line), or set to $null for a full tenant scan
        - Internal name of the column to remove ($columnName)
        - Lists to ignore (default: 'Site Pages')
        - Throttling delay values if needed
        
        Throttling Protection:
        - DelayBetweenLists: 500ms (default)
        - DelayBetweenSites: 1000ms (default)
        - MaxRetryAttempts: 5
        - BaseRetryDelay: 5000ms with exponential backoff
        
        Logging:
        All operations are logged to a timestamped file in %TEMP% directory
        
    .LINK
        https://pnp.github.io/powershell/

    .OUTPUTS
        Log file in %TEMP% directory named: Remove_SPO_Column_[timestamp]_logfile.log
        Console output showing progress and results
    #>

# =================================================================================================
# USER CONFIGURATION - Update the variables in this section
# =================================================================================================

# --- Script Parameters ---
param(
    [switch]$Verbose = $false  # Add -Verbose to see detailed compliance flag information
)

# --- Tenant and App Registration Details ---
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"                 # This is your Entra App ID
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"                # This is your Tenant ID
$thumbprint = "216f5dd7327719bc8cf15ff3c077adf59ace0c23"        # This is certificate thumbprint
$tenantAdminUrl = "https://m365cpi13246019-admin.sharepoint.com"        # Tenant admin URL (required for full tenant scan)

# --- Input File Path ---
$sitelist = '' # Path to file containing site URLs (one per line). Set to $null or '' to scan all tenant sites.

# --- Lists to Ignore ---
$ignoreListNames = @('Site Pages') # Lists to skip during processing

# --- Column to Remove ---
$columnName = "yourcolumnname"             # Internal name of the column to remove from all lists/libraries

# --- Logging Configuration ---
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss" # Current date and time for unique file naming
$script:LogFilePath = "$env:TEMP\" + 'Remove_SPO_Column_' + $date + '_' + "logfile.log" # Define log file path
$script:EnableLogging = $true

# --- Throttling Protection Settings ---
$DelayBetweenLists = 500        # Milliseconds delay between processing each list (default: 500ms)
$DelayBetweenSites = 1000       # Milliseconds delay between processing each site (default: 1000ms)
$MaxRetryAttempts = 5           # Maximum number of retry attempts for throttled requests
$BaseRetryDelay = 5000          # Base delay in milliseconds for exponential backoff (default: 5 seconds)

# =================================================================================================
# END OF USER CONFIGURATION
# =================================================================================================

#region LOGGING FUNCTIONS
# =====================================================================================
# Logging functionality to capture console output to file
# =====================================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'VERBOSE', 'SUCCESS')]
        [string]$Level = 'INFO',
        
        [Parameter()]
        [string]$LogFile = $script:LogFilePath
    )
    
    if (-not $script:EnableLogging -or [string]::IsNullOrEmpty($LogFile)) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8
    }
    catch {
        # If logging fails, don't break the script
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
    }
}

function Write-VerboseAndLog {
    param([string]$Message)
    Write-Verbose $Message
    Write-Log -Message $Message -Level 'VERBOSE'
}

function Write-WarningAndLog {
    param([string]$Message)
    Write-Warning $Message
    Write-Log -Message $Message -Level 'WARNING'
}

function Write-HostAndLog {
    param(
        [string]$Object,
        [string]$ForegroundColor,
        [switch]$NoNewline
    )
    
    if ($ForegroundColor) {
        if ($NoNewline) {
            Write-Host $Object -ForegroundColor $ForegroundColor -NoNewline
        }
        else {
            Write-Host $Object -ForegroundColor $ForegroundColor
        }
    }
    else {
        if ($NoNewline) {
            Write-Host $Object -NoNewline
        }
        else {
            Write-Host $Object
        }
    }
    
    # Only log if the message is not empty
    if (-not [string]::IsNullOrWhiteSpace($Object)) {
        # Determine log level based on color
        $logLevel = switch ($ForegroundColor) {
            'Red' { 'WARNING' }  # Red items are findings, not errors (locked items found)
            'Yellow' { 'WARNING' }
            'Green' { 'SUCCESS' }
            default { 'INFO' }
        }
        
        Write-Log -Message $Object -Level $logLevel
    }
}

#endregion LOGGING FUNCTIONS

#region THROTTLING FUNCTIONS
# =====================================================================================
# Throttling protection functionality to handle SharePoint Online throttling
# =====================================================================================

function Invoke-WithThrottlingRetry {
    <#
    .SYNOPSIS
    Executes a script block with automatic retry logic for throttling (429/503 responses)
    
    .DESCRIPTION
    This function implements the Microsoft recommended approach for handling SharePoint
    throttling by honoring Retry-After headers and using exponential backoff.
    
    .PARAMETER ScriptBlock
    The script block to execute
    
    .PARAMETER MaxRetries
    Maximum number of retry attempts (default: from configuration)
    
    .PARAMETER Description
    Description of the operation being performed (for logging)
    #>
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,
        
        [Parameter()]
        [int]$MaxRetries = $script:MaxRetryAttempts,
        
        [Parameter()]
        [string]$Description = "Operation"
    )
    
    $attempt = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $attempt -lt $MaxRetries) {
        $attempt++
        
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $statusCode = $null
            $retryAfter = $null
            
            # Try to extract status code from different exception types
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                
                # Try to get Retry-After header
                if ($_.Exception.Response.Headers) {
                    $retryAfterHeader = $_.Exception.Response.Headers["Retry-After"]
                    if ($retryAfterHeader) {
                        $retryAfter = [int]$retryAfterHeader
                    }
                }
            }
            elseif ($_.Exception.Message -match "429|503") {
                # Try to extract from error message
                if ($_.Exception.Message -match "429") {
                    $statusCode = 429
                }
                elseif ($_.Exception.Message -match "503") {
                    $statusCode = 503
                }
            }
            
            # Check if this is a throttling error (429 or 503)
            if ($statusCode -eq 429 -or $statusCode -eq 503) {
                if ($attempt -lt $MaxRetries) {
                    # Calculate delay: use Retry-After if available, otherwise exponential backoff
                    $delaySeconds = if ($retryAfter) {
                        $retryAfter
                    }
                    else {
                        # Exponential backoff: BaseDelay * 2^(attempt-1)
                        ($script:BaseRetryDelay / 1000) * [Math]::Pow(2, $attempt - 1)
                    }
                    
                    Write-WarningAndLog "$Description - Throttled (HTTP $statusCode). Waiting $delaySeconds seconds before retry $attempt of $MaxRetries..."
                    Start-Sleep -Seconds $delaySeconds
                }
                else {
                    Write-HostAndLog "✗ $Description - Max retry attempts reached after throttling" -ForegroundColor Red
                    throw
                }
            }
            else {
                # Non-throttling error, don't retry
                throw
            }
        }
    }
    
    return $result
}

function Start-ThrottleDelay {
    <#
    .SYNOPSIS
    Introduces a delay to avoid overwhelming SharePoint with requests
    
    .PARAMETER DelayMilliseconds
    The delay in milliseconds
    
    .PARAMETER Description
    Description of what we're delaying (for verbose logging)
    #>
    param(
        [Parameter(Mandatory)]
        [int]$DelayMilliseconds,
        
        [Parameter()]
        [string]$Description = "throttle protection"
    )
    
    if ($DelayMilliseconds -gt 0) {
        Write-VerboseAndLog "Applying $DelayMilliseconds ms delay for $Description"
        Start-Sleep -Milliseconds $DelayMilliseconds
    }
}

#endregion THROTTLING FUNCTIONS

# Function to check if a column exists in a list or library
function Test-ColumnExistsInList {
    param(
        [string]$ListTitle
    )
    
    try {
        $field = Invoke-WithThrottlingRetry -Description "Check for column '$columnName' in list '$ListTitle'" -ScriptBlock {
            Get-PnPField -List $ListTitle -Identity $columnName -ErrorAction SilentlyContinue
        }
        return ($null -ne $field)
    }
    catch {
        Write-VerboseAndLog "Could not check column existence in list '$ListTitle': $($_.Exception.Message)"
        return $false
    }
}

# Function to remove a column from a list or library
function Remove-ColumnFromList {
    param(
        [string]$ListTitle
    )
    
    try {
        Invoke-WithThrottlingRetry -Description "Remove column '$columnName' from list '$ListTitle'" -ScriptBlock {
            Remove-PnPField -List $ListTitle -Identity $columnName -Force -ErrorAction Stop
        }
        Write-HostAndLog "  ✓ Successfully removed column '$columnName' from list '$ListTitle'" -ForegroundColor Green
        return $true
    }
    catch {
        Write-HostAndLog "  ✗ Failed to remove column '$columnName' from list '$ListTitle': $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Resolve site URLs: read from file, or perform a full tenant scan
if (-not [string]::IsNullOrWhiteSpace($sitelist)) {
    # --- File-based mode ---
    if (-not (Test-Path $sitelist)) {
        Write-HostAndLog "Input file not found: $sitelist" -ForegroundColor Red
        exit 1
    }
    $siteUrls = Get-Content $sitelist | Where-Object { $_.Trim() -ne "" }
    $scanMode = "File: $sitelist"
}
else {
    # --- Full tenant scan mode ---
    Write-HostAndLog "No input file specified - performing full tenant scan..." -ForegroundColor Cyan
    Write-Log "No input file specified - connecting to tenant admin for full site list" -Level 'INFO'
    
    try {
        Invoke-WithThrottlingRetry -Description "Connect to tenant admin" -ScriptBlock {
            Connect-PnPOnline -Url $tenantAdminUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        
        $allSites = Invoke-WithThrottlingRetry -Description "Get all tenant sites" -ScriptBlock {
            Get-PnPTenantSite -IncludeOneDriveSites:$false | Select-Object -ExpandProperty Url
        }
        
        Disconnect-PnPOnline
        $siteUrls = $allSites
        $scanMode = "Full tenant scan ($($siteUrls.Count) sites discovered)"
        Write-HostAndLog "Tenant scan complete - discovered $($siteUrls.Count) sites" -ForegroundColor Cyan
        Write-Log "Tenant scan complete - discovered $($siteUrls.Count) sites" -Level 'INFO'
    }
    catch {
        Write-HostAndLog "Failed to retrieve tenant site list: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Failed to retrieve tenant site list: $($_.Exception.Message)" -Level 'ERROR'
        exit 1
    }
}

# Initialize logging
Write-Log "=== REMOVE SPO COLUMN SCRIPT STARTED ===" -Level 'INFO'
Write-Log "Log file: $script:LogFilePath" -Level 'INFO'
Write-Log "Scan mode: $scanMode" -Level 'INFO'
Write-Log "Column to remove: $columnName" -Level 'INFO'
Write-Log "Tenant: $tenant" -Level 'INFO'
Write-Log "AppID: $appID" -Level 'INFO'

Write-HostAndLog "Starting to process $($siteUrls.Count) SharePoint sites..." -ForegroundColor Cyan
Write-HostAndLog "Scan mode: $scanMode" -ForegroundColor Cyan
Write-HostAndLog "Column to remove: $columnName" -ForegroundColor Cyan
Write-HostAndLog "Log file location: $script:LogFilePath" -ForegroundColor Gray
Write-HostAndLog "" -ForegroundColor Gray
Write-HostAndLog "Throttling Protection: Enabled" -ForegroundColor Cyan
Write-HostAndLog "  - Delay between lists: $DelayBetweenLists ms" -ForegroundColor Gray
Write-HostAndLog "  - Delay between sites: $DelayBetweenSites ms" -ForegroundColor Gray
Write-HostAndLog "  - Max retry attempts: $MaxRetryAttempts" -ForegroundColor Gray
Write-HostAndLog "  - Base retry delay: $BaseRetryDelay ms" -ForegroundColor Gray

$totalRemovedColumns = 0
$totalListsWithColumn = 0
$totalProcessedSites = 0
$totalProcessedLists = 0
$totalSiteCount = $siteUrls.Count
$siteCounter = 0

foreach ($siteUrl in $siteUrls) {
    $siteUrl = $siteUrl.Trim()
    $siteCounter++
    Write-HostAndLog "`n[$siteCounter/$totalSiteCount] Processing site: $siteUrl" -ForegroundColor Yellow
    
    try {
        # Connect to the current SharePoint site with throttling retry
        Invoke-WithThrottlingRetry -Description "Connect to site $siteUrl" -ScriptBlock {
            Connect-PnPOnline -Url $siteUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        $totalProcessedSites++
        Write-Log "Successfully connected to site: $siteUrl" -Level 'SUCCESS'
        
        # Get all lists in the site with throttling retry
        $lists = Invoke-WithThrottlingRetry -Description "Get lists from site $siteUrl" -ScriptBlock {
            Get-PnPList | Where-Object { 
                $_.Hidden -eq $false -and 
                $_.Title -notin $ignoreListNames 
            }
        }
        
        Write-HostAndLog "Found $($lists.Count) non-hidden lists/libraries in this site" -ForegroundColor Cyan
        
        $listCounter = 0
        foreach ($list in $lists) {
            $listCounter++
            Write-HostAndLog "  Processing list $listCounter of $($lists.Count): $($list.Title)" -ForegroundColor White
            $totalProcessedLists++
            
            try {
                # Check if the column exists in this list/library
                if (Test-ColumnExistsInList -ListTitle $list.Title) {
                    $totalListsWithColumn++
                    Write-HostAndLog "    Column '$columnName' found - removing..." -ForegroundColor Yellow
                    
                    if (Remove-ColumnFromList -ListTitle $list.Title) {
                        $totalRemovedColumns++
                    }
                }
                else {
                    Write-VerboseAndLog "    Column '$columnName' not found in list '$($list.Title)' - skipping"
                }
                
                # Throttle protection: delay between lists
                if ($listCounter -lt $lists.Count) {
                    Start-ThrottleDelay -DelayMilliseconds $DelayBetweenLists -Description "between lists"
                }
            }
            catch {
                Write-HostAndLog "    Error processing list '$($list.Title)': $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Disconnect from current site
        Disconnect-PnPOnline
        Write-Log "Disconnected from site: $siteUrl" -Level 'INFO'
        
        # Throttle protection: delay between sites
        Start-ThrottleDelay -DelayMilliseconds $DelayBetweenSites -Description "between sites"
    }
    catch {
        Write-HostAndLog "Error connecting to or processing site '$siteUrl': $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Final summary
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "PROCESSING COMPLETE" -ForegroundColor Cyan
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "Scan mode: $scanMode" -ForegroundColor Cyan
Write-HostAndLog "Column targeted: $columnName" -ForegroundColor Cyan
Write-HostAndLog "Sites processed: $totalProcessedSites of $($siteUrls.Count)" -ForegroundColor Yellow
Write-HostAndLog "Lists/libraries processed: $totalProcessedLists" -ForegroundColor Yellow
Write-HostAndLog "Lists/libraries where column was found: $totalListsWithColumn" -ForegroundColor Yellow
Write-HostAndLog "Total column removals succeeded: $totalRemovedColumns" -ForegroundColor Green
Write-HostAndLog "=================================" -ForegroundColor Cyan
Write-HostAndLog "Log file saved to: $script:LogFilePath" -ForegroundColor Gray

# Final log entries
Write-Log "=== PROCESSING SUMMARY ===" -Level 'INFO'
Write-Log "Scan mode: $scanMode" -Level 'INFO'
Write-Log "Column targeted: $columnName" -Level 'INFO'
Write-Log "Sites processed: $totalProcessedSites of $($siteUrls.Count)" -Level 'INFO'
Write-Log "Lists/libraries processed: $totalProcessedLists" -Level 'INFO'
Write-Log "Lists/libraries where column was found: $totalListsWithColumn" -Level 'INFO'
Write-Log "Total column removals succeeded: $totalRemovedColumns" -Level 'SUCCESS'
Write-Log "Throttling protection settings: Lists=$DelayBetweenLists ms, Sites=$DelayBetweenSites ms, MaxRetries=$MaxRetryAttempts" -Level 'INFO'
Write-Log "=== REMOVE SPO COLUMN SCRIPT COMPLETED ===" -Level 'INFO'
