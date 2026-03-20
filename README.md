# Remove-SPOColumn.ps1

A PowerShell script that removes a specified column from all lists and libraries across multiple SharePoint Online sites. Supports both targeted site-list mode and full tenant scan mode, with built-in throttling protection, exponential backoff retry logic, and detailed logging.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Authentication Setup](#authentication-setup)
- [Configuration](#configuration)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Throttling Protection](#throttling-protection)
- [Logging](#logging)
- [Output and Summary](#output-and-summary)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)

---

## Overview

This script connects to one or more SharePoint Online sites and removes a specified column (field) from every list and library where it exists. It is designed for bulk column cleanup operations at scale and handles both site columns and list-level fields.

**Key capabilities:**
- Processes all non-hidden lists and libraries in each site
- Skips lists where the column is not present (no false errors)
- Supports targeting a specific list of sites via a text file, or scanning all sites in the tenant automatically
- Handles SharePoint throttling (HTTP 429/503) with Retry-After header support and exponential backoff
- Logs all activity to a timestamped log file

---

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | PowerShell 7+ recommended (compatible with Windows PowerShell 5.1) |
| PnP.PowerShell | Install via `Install-Module PnP.PowerShell` |
| Entra App Registration | App registration with certificate-based authentication |
| Permissions | SharePoint:Sites.FullControl.All` |

Install the PnP.PowerShell module if not already present:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

---

## Authentication Setup

The script uses **certificate-based authentication** via an Entra (Azure AD) App Registration. No user credentials are stored or prompted.

### Steps to configure authentication:

1. **Create an App Registration** in the [Entra Admin Center](https://entra.microsoft.com)
2. **Grant API permissions:**
   - `SharePoint > Sites.FullControl.All` (Application permission)
   - Grant admin consent 
3. **Upload or generate a certificate** on the App Registration under *Certificates & Secrets*
4. **Install the certificate** to the local machine's certificate store (or note the thumbprint if already installed)
5. **Note the following values** for the configuration section:
   - Application (client) ID
   - Directory (tenant) ID
   - Certificate thumbprint

---

## Configuration

Before running the script, update the **USER CONFIGURATION** section near the top of the file:

```powershell
# --- Tenant and App Registration Details ---
$appID        = "your-entra-app-id"
$tenant       = "your-tenant-id"
$thumbprint   = "your-certificate-thumbprint"
$tenantAdminUrl = "https://your-tenant-admin.sharepoint.com"

# --- Input File Path ---
# Provide a path to a text file with one site URL per line,
# OR leave empty ('') to scan all sites in the tenant.
$sitelist = ''

# --- Lists to Ignore ---
$ignoreListNames = @('Site Pages')

# --- Column to Remove ---
$columnName = "YourColumnInternalName"
```

### Configuration Reference

| Variable | Description | Example |
|---|---|---|
| `$appID` | Entra App Registration client ID | `"1e488dc4-..."` |
| `$tenant` | Entra tenant ID (GUID) | `"9cfc42cb-..."` |
| `$thumbprint` | Certificate thumbprint | `"216f5dd7..."` |
| `$tenantAdminUrl` | SharePoint admin center URL | `"https://contoso-admin.sharepoint.com"` |
| `$sitelist` | Path to sites file, or `''` for full tenant scan | `"C:\sites.txt"` or `''` |
| `$ignoreListNames` | Array of list names to skip | `@('Site Pages', 'Style Library')` |
| `$columnName` | **Internal name** of the column to remove | `"YourColumnInternalName"` |
| `$DelayBetweenLists` | Milliseconds to wait between each list | `500` |
| `$DelayBetweenSites` | Milliseconds to wait between each site | `1000` |
| `$MaxRetryAttempts` | Max retries on throttled requests | `5` |
| `$BaseRetryDelay` | Base delay (ms) for exponential backoff | `5000` |

> **Important:** Use the column's **internal name**, not its display name. To find the internal name, navigate to the column settings in SharePoint and check the URL for the `Field=` parameter, or use PnP PowerShell: `Get-PnPField | Select-Object Title, InternalName`

---

## Usage

### Basic run (default settings)

```powershell
.\Remove-SPOColumn.ps1
```

### Run with verbose output

Displays additional diagnostic details for each list processed, including lists where the column was not found.

```powershell
.\Remove-SPOColumn.ps1 -Verbose
```

### Site list file format

If using `$sitelist`, create a plain text file with one SharePoint site URL per line:

```
https://contoso.sharepoint.com/sites/HR
https://contoso.sharepoint.com/sites/Finance
https://contoso.sharepoint.com/teams/ProjectAlpha
```

---

## How It Works

```
Script Start
    │
    ├─► If $sitelist is set → Load site URLs from file
    └─► If $sitelist is empty → Connect to Tenant Admin → Get-PnPTenantSite → Disconnect
    
For each Site URL:
    ├─► Connect-PnPOnline (certificate auth)
    ├─► Get-PnPList (non-hidden, not in ignore list)
    │
    └─► For each List:
            ├─► Get-PnPField → Does column exist?
            │       ├─► No  → Skip (verbose log only)
            │       └─► Yes → Remove-PnPField -Force
            │                   ├─► Success → Log green ✓
            │                   └─► Failure → Log red ✗
            └─► Apply inter-list delay
    
    ├─► Disconnect-PnPOnline
    └─► Apply inter-site delay
    
Print summary → Write log file
```

---

## Throttling Protection

SharePoint Online enforces request throttling. This script implements Microsoft's recommended approach:

| Protection | Behavior |
|---|---|
| **HTTP 429 / 503 detection** | Catches throttling responses from both exception codes and error message text |
| **Retry-After header** | If SharePoint returns a `Retry-After` header, the script waits exactly that duration |
| **Exponential backoff** | If no `Retry-After` header, delay is `BaseRetryDelay × 2^(attempt-1)` |
| **Max retries** | After `$MaxRetryAttempts` (default: 5), the operation is logged as failed and processing continues |
| **Inter-list delay** | Configurable delay between processing each list (default: 500ms) |
| **Inter-site delay** | Configurable delay between processing each site (default: 1000ms) |

The exponential backoff sequence with default settings (5s base):

| Attempt | Delay |
|---|---|
| 1 | 5 seconds |
| 2 | 10 seconds |
| 3 | 20 seconds |
| 4 | 40 seconds |
| 5 | 80 seconds |

---

## Logging

All output is written to both the console and a timestamped log file.

**Log file location:**
```
%TEMP%\Remove_SPO_Column_yyyy-MM-dd_HH-mm-ss_logfile.log
```

**Log levels:**

| Level | Used for |
|---|---|
| `INFO` | General progress messages |
| `SUCCESS` | Successful column removals and connections |
| `WARNING` | Throttling retries, non-critical issues |
| `ERROR` | Failures to connect or retrieve site list |
| `VERBOSE` | Detailed per-list diagnostics (when `-Verbose` is used) |

---

## Output and Summary

At the end of the run, the script prints a summary to the console and writes it to the log file:

```
=================================
PROCESSING COMPLETE
=================================
Scan mode:                      File: C:\sites.txt
Column targeted:                YourColumnInternalName
Sites processed:                42 of 42
Lists/libraries processed:      387
Lists/libraries with column:    23
Total column removals succeeded: 23
=================================
Log file saved to: C:\Users\...\AppData\Local\Temp\Remove_SPO_Column_..._logfile.log
```

---

## Examples

### Remove a column from a specific set of sites

1. Create `sites.txt`:
   ```
   https://contoso.sharepoint.com/sites/Legal
   https://contoso.sharepoint.com/sites/Compliance
   ```
2. Set `$sitelist = "C:\sites.txt"` and `$columnName = "YourColumnInternalName"` in the script
3. Run:
   ```powershell
   .\Remove-SPOColumn.ps1
   ```

### Remove a column from all sites in the tenant

1. Set `$sitelist = ''` and `$tenantAdminUrl = "https://contoso-admin.sharepoint.com"`
2. Run:
   ```powershell
   .\Remove-SPOColumn.ps1 -Verbose
   ```

---

## Troubleshooting

| Issue | Resolution |
|---|---|
| `Connect-PnPOnline` fails | Verify `$appID`, `$tenant`, `$thumbprint` are correct and the certificate is installed in the current user's certificate store |
| Column not found / not removed | Ensure `$columnName` is the **internal name**, not the display name |
| `Access Denied` on some sites | The app registration may not have access; check API permissions and admin consent |
| Repeated throttling | Increase `$DelayBetweenLists`, `$DelayBetweenSites`, or `$BaseRetryDelay` to reduce request rate |
| Full tenant scan returns 0 sites | Confirm `$tenantAdminUrl` is correct and the app has `Sites.FullControl.All` or `SharePoint > Sites.ReadWrite.All` with tenant-wide access |
| Log file not created | Check write permissions on `%TEMP%`; set `$script:EnableLogging = $false` to disable if needed |

---

## Author

**Mike Lee**  
Date: March 20, 2026

