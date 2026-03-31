<#
.SYNOPSIS
    Finds SharePoint Online FileDownloaded events performed via the Seclore-O365
    client app, tenant-wide, using the Unified Audit Log.

.DESCRIPTION
    Queries Search-UnifiedAuditLog for the SharePoint FileDownloaded operation
    across the entire tenant and filters results to only those where the
    AuditData AppAccessContext contains ClientAppName = "Seclore-O365".

.PARAMETER StartDate
    Start of the audit window (UTC). Defaults to 3 days ago.

.PARAMETER EndDate
    End of the audit window (UTC). Defaults to now.

.PARAMETER OutputPath
    Path for the exported CSV. Defaults to .\SPO-SecloreDownloads_<timestamp>.csv

.PARAMETER ResultSize
    Max records returned per Search-UnifiedAuditLog call. Max allowed is 5000.

.EXAMPLE
    .\find-seclore.ps1 -StartDate (Get-Date).AddDays(-30)

.NOTES
    Author  : Mike Lee
    Created : 3/27/2026
    Version : 3.0
#>

[CmdletBinding()]
param (
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

function Get-AuditEvents {
    <#
    .SYNOPSIS  Pages through Search-UnifiedAuditLog tenant-wide (no site scope).
    #>
    param (
        [datetime] $Start,
        [datetime] $End,
        [string[]] $Operations,
        [int]      $PageSize,
        [string]   $FreeText
    )

    $allRecords = [System.Collections.Generic.List[PSObject]]::new()
    $sessionId = "SPOSecloreAudit-$(New-Guid)"
    $page = 1

    Write-Verbose "  Querying UAL tenant-wide for: $($Operations -join ', ')"

    do {
        $results = Search-UnifiedAuditLog `
            -StartDate       $Start `
            -EndDate         $End `
            -Operations      $Operations `
            -FreeText        $FreeText `
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

    # AppAccessContext can be a nested JSON string in tenant-wide UAL results
    $appCtx = $audit.AppAccessContext
    if ($appCtx -is [string] -and $appCtx) {
        try { $appCtx = $appCtx | ConvertFrom-Json -ErrorAction Stop } catch { $appCtx = $null }
    }

    [PSCustomObject]@{
        DateTime      = $Record.CreationDate
        PerformedBy   = $Record.UserIds
        Operation     = $Record.Operations
        ItemType      = $audit.ItemType
        FileName      = $audit.SourceFileName
        RelativePath  = $relPath
        SiteUrl       = $audit.SiteUrl
        ClientAppName = $appCtx.ClientAppName
        ClientAppId   = $appCtx.ClientAppId
        ClientIP      = $audit.ClientIP
        UserAgent     = $audit.UserAgent
    }
}

#endregion

#region ── Main ─────────────────────────────────────────────────────────────────

Write-Host "SPO Seclore Download Audit" -ForegroundColor Cyan
Write-Host "  Scope    : Tenant-wide" -ForegroundColor Cyan
Write-Host "  Window   : $($StartDate.ToString('u'))  ->  $($EndDate.ToString('u'))" -ForegroundColor Cyan
Write-Host "  Filter   : Operation=FileDownloaded, ClientAppName=Seclore-O365" -ForegroundColor Cyan
Write-Host ""

$allResults = [System.Collections.Generic.List[PSObject]]::new()

Write-Progress -Activity "Querying Unified Audit Log" -Status "Retrieving tenant-wide FileDownloaded events..."

try {
    $records = Get-AuditEvents `
        -Start      $StartDate `
        -End        $EndDate `
        -Operations $PermissionOperations `
        -FreeText   'Seclore-O365' `
        -PageSize   $ResultSize

    Write-Progress -Activity "Querying Unified Audit Log" -Status "Filtering for Seclore-O365..."

    if ($records -and $records.Count -gt 0) {
        foreach ($rec in $records) {
            # Pre-filter on raw string before JSON parse (handles nested JSON string edge case)
            if ($rec.AuditData -notlike '*Seclore-O365*') { continue }
            try { $audit = $rec.AuditData | ConvertFrom-Json -ErrorAction Stop }
            catch { continue }
            # Resolve AppAccessContext whether it is an object or a nested JSON string
            $appCtx = $audit.AppAccessContext
            if ($appCtx -is [string] -and $appCtx) {
                try { $appCtx = $appCtx | ConvertFrom-Json -ErrorAction Stop } catch { $appCtx = $null }
            }
            if ($appCtx.ClientAppName -ne 'Seclore-O365') { continue }
            $allResults.Add((ConvertTo-FlatRecord -Record $rec))
        }
        Write-Host "  Total FileDownloaded events  : $($records.Count)" -ForegroundColor Gray
        Write-Host "  Seclore-O365 matches         : $($allResults.Count)" -ForegroundColor $(if ($allResults.Count -gt 0) { 'Green' } else { 'Yellow' })
    }
    else {
        Write-Host "  No FileDownloaded events found in the specified window." -ForegroundColor Yellow
    }
}
catch {
    Write-Warning "  [!] Error querying audit log: $_"
}

Write-Progress -Activity "Querying Unified Audit Log" -Completed

#endregion

#region -------------------------- Output ---------------------------------------------------------

if ($allResults.Count -eq 0) {
    Write-Host "`nNo Seclore-O365 FileDownloaded events found in the specified window." -ForegroundColor Yellow
}
else {
    $allResults |
    Sort-Object DateTime -Descending |
    Select-Object DateTime, PerformedBy, Operation, ItemType, FileName, RelativePath,
    SiteUrl, ClientAppName, ClientAppId, ClientIP, UserAgent |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nResults  : $($allResults.Count) Seclore-O365 FileDownloaded event(s)" -ForegroundColor Cyan
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
