<#
.SYNOPSIS
    Finds SharePoint Online FileDownloaded events performed via the Seclore-o365
    client app, across a list of site collections, using the Unified Audit Log.

.DESCRIPTION
    Queries Search-UnifiedAuditLog for the SharePoint FileDownloaded operation
    and filters results to only those where the AuditData AppAccessContext
    contains ClientAppName = "Seclore-o365".

.PARAMETER SiteListPath
    Path to a text file containing one SPO site URL per line.

.PARAMETER StartDate
    Start of the audit window (UTC). Defaults to 3 days ago.

.PARAMETER EndDate
    End of the audit window (UTC). Defaults to now.

.PARAMETER OutputPath
    Path for the exported CSV. Defaults to .\SPO-SecloreDownloads_<timestamp>.csv

.PARAMETER ResultSize
    Max records returned per Search-UnifiedAuditLog call. Max allowed is 5000.

.EXAMPLE
    .\find-seclore.ps1 -SiteListPath .\sites.txt -StartDate (Get-Date).AddDays(-30)

.NOTES
    Author  : Mike Lee / Mariel Williams
    Created : 3/27/2026
    Version : 2.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SiteListPath = "C:\temp\SPOSiteList.txt",

    [Parameter()]
    [datetime]$StartDate = (Get-Date).ToUniversalTime().AddDays(-3),

    [Parameter()]
    [datetime]$EndDate = (Get-Date).ToUniversalTime(),

    [Parameter()]
    [string]$OutputPath = ".\SPO-SecloreDownloads_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [ValidateRange(1, 5000)]
    [int]$ResultSize = 5000
)

#region ── Prerequisites ────────────────────────────────────────────────────────

# Verify ExchangeOnlineManagement is available (provides Search-UnifiedAuditLog)
if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

#endregion

#region ── Configuration ────────────────────────────────────────────────────────



# Target operation: SharePoint file downloads only
$PermissionOperations = @(
    'FileDownloaded'
)

#endregion

#region ── Helpers ──────────────────────────────────────────────────────────────

function Get-AuditEventsForSite {
    <#
    .SYNOPSIS  Pages through Search-UnifiedAuditLog for a single site URL.
    #>
    param (
        [string]   $SiteUrl,
        [datetime] $Start,
        [datetime] $End,
        [string[]] $Operations,
        [int]      $PageSize
    )

    $allRecords = [System.Collections.Generic.List[PSObject]]::new()
    $sessionId = "SPOPermAudit-$(New-Guid)"
    $page = 1

    Write-Verbose "  Querying UAL for: $SiteUrl"

    do {
        $results = Search-UnifiedAuditLog `
            -StartDate       $Start `
            -EndDate         $End `
            -Operations      $Operations `
            -ObjectIds       "$SiteUrl*" `
            -SessionId       $sessionId `
            -SessionCommand  ReturnLargeSet `
            -ResultSize      $PageSize `
            -ErrorAction     Stop

        if ($results) {
            foreach ($r in $results) { $allRecords.Add($r) }
            Write-Verbose "    Page $page - retrieved $($results.Count) records (running total: $($allRecords.Count))"
            $page++
        }
    } while ($results -and $results.Count -eq $PageSize)

    return $allRecords
}

function ConvertTo-FlatRecord {
    <#
    .SYNOPSIS  Flattens a UAL FileDownloaded record into a clean, admin-readable object.
    #>
    param ([PSObject]$Record)

    try {
        $audit = $Record.AuditData | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $audit = $null
    }

    # Relative path: SourceRelativeUrl is cleanest; fall back to stripping the site URL from ObjectId
    $relPath = ''
    if ($audit.SourceRelativeUrl) {
        $relPath = $audit.SourceRelativeUrl
    }
    elseif ($audit.ObjectId -and $audit.SiteUrl) {
        $stripped = $audit.ObjectId -replace [regex]::Escape($audit.SiteUrl.TrimEnd('/')), ''
        $relPath = if ($stripped -match '^[/\\]?$') { '(site root)' } else { $stripped.TrimStart('/') }
    }

    [PSCustomObject]@{
        DateTime      = $Record.CreationDate
        PerformedBy   = $Record.UserIds
        Operation     = $Record.Operations
        ItemType      = $audit.ItemType
        FileName      = $audit.SourceFileName
        RelativePath  = $relPath
        SiteUrl       = $audit.SiteUrl
        ClientAppName = $audit.AppAccessContext.ClientAppName
        ClientAppId   = $audit.AppAccessContext.ClientAppId
        ClientIP      = $audit.ClientIP
        UserAgent     = $audit.UserAgent
    }
}

#endregion

#region ── Main ─────────────────────────────────────────────────────────────────


# Load site list - skip blank lines and comment lines
$sites = Get-Content -Path $SiteListPath |
Where-Object { $_ -match 'https?://' } |
ForEach-Object { $_.Trim().TrimEnd('/') } |
Select-Object -Unique

if (-not $sites) {
    throw "No valid SPO URLs found in '$SiteListPath'. Each line should contain a URL starting with https://"
}

Write-Host "SPO Seclore Download Audit" -ForegroundColor Cyan
Write-Host "  Sites    : $($sites.Count)" -ForegroundColor Cyan
Write-Host "  Window   : $($StartDate.ToString('u'))  ->  $($EndDate.ToString('u'))" -ForegroundColor Cyan
Write-Host "  Filter   : Operation=FileDownloaded, ClientAppName=Seclore-o365" -ForegroundColor Cyan
Write-Host ""

$allResults = [System.Collections.Generic.List[PSObject]]::new()
$siteIndex = 0

foreach ($site in $sites) {
    $siteIndex++
    Write-Progress -Activity "Querying Unified Audit Log" `
        -Status    "[$siteIndex/$($sites.Count)] $site" `
        -PercentComplete (($siteIndex / $sites.Count) * 100)

    try {
        $records = Get-AuditEventsForSite `
            -SiteUrl    $site `
            -Start      $StartDate `
            -End        $EndDate `
            -Operations $PermissionOperations `
            -PageSize   $ResultSize

        if ($records -and $records.Count -gt 0) {
            $secloreCount = 0
            foreach ($rec in $records) {
                try { $audit = $rec.AuditData | ConvertFrom-Json -ErrorAction Stop }
                catch { continue }
                if ($audit.AppAccessContext.ClientAppName -ne 'Seclore-o365') { continue }
                $allResults.Add((ConvertTo-FlatRecord -Record $rec))
                $secloreCount++
            }
            if ($secloreCount -gt 0) {
                Write-Host "  [+] $site  -  $secloreCount Seclore download(s) (of $($records.Count) total FileDownloaded)" -ForegroundColor Green
            }
            else {
                Write-Host "  [ ] $site  -  no Seclore downloads (of $($records.Count) total FileDownloaded)" -ForegroundColor Gray
            }
        }
        else {
            Write-Host "  [ ] $site  -  no events found" -ForegroundColor Gray
        }
    }
    catch {
        Write-Warning "  [!] Error querying '$site': $_"
    }
}

Write-Progress -Activity "Querying Unified Audit Log" -Completed

#endregion

#region ── Output ───────────────────────────────────────────────────────────────

if ($allResults.Count -eq 0) {
    Write-Host "`nNo Seclore-o365 FileDownloaded events found across any sites in the specified window." -ForegroundColor Yellow
}
else {
    $allResults |
    Sort-Object DateTime -Descending |
    Select-Object DateTime, PerformedBy, Operation, ItemType, FileName, RelativePath,
    SiteUrl, ClientAppName, ClientAppId, ClientIP, UserAgent |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nResults  : $($allResults.Count) Seclore-o365 FileDownloaded event(s)" -ForegroundColor Cyan
    Write-Host "Exported : $OutputPath" -ForegroundColor Green

    # Summary by site
    Write-Host "`n── Downloads per site ──────────────────────────────────" -ForegroundColor Cyan
    $allResults |
    Group-Object SiteUrl |
    Sort-Object Count -Descending |
    Format-Table @{L = 'SiteUrl'; E = { $_.Name }; W = 60 }, Count -AutoSize

    # Who downloaded files
    Write-Host "── Downloads by user ───────────────────────────────────" -ForegroundColor Cyan
    $allResults |
    Group-Object PerformedBy |
    Sort-Object Count -Descending |
    Format-Table @{L = 'PerformedBy'; E = { $_.Name }; W = 50 }, Count -AutoSize
}

#endregion
