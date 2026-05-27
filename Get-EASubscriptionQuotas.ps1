<#
.SYNOPSIS
    Lists all EA subscriptions and their resource quotas, exporting results to an Excel file.

.DESCRIPTION
    This script enumerates all Azure subscriptions accessible to the current user,
    retrieves quota usage for Compute, Network, and Storage providers across all regions
    where resources are deployed, and exports the data to a formatted Excel workbook.

    Requires:
    - Windows PowerShell 5.1
    - Azure CLI (az) installed and logged in (run 'az login' first)
    - Internet access (to install the ImportExcel module if not already present)

.PARAMETER OutputPath
    Path for the output Excel file. Defaults to EA-Subscription-Quotas-<date>.xlsx in the current directory.

.PARAMETER SubscriptionFilter
    Optional. Filter subscriptions by name pattern (supports wildcards). Default: "*" (all).

.PARAMETER MaxRetries
    Maximum number of retries for throttled API calls. Default: 3.

.EXAMPLE
    .\Get-EASubscriptionQuotas.ps1
    .\Get-EASubscriptionQuotas.ps1 -OutputPath "C:\Reports\quotas.xlsx"
    .\Get-EASubscriptionQuotas.ps1 -SubscriptionFilter "Prod*"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path (Get-Location) "EA-Subscription-Quotas-$(Get-Date -Format 'yyyy-MM-dd').xlsx"),

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionFilter = "*",

    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 3
)

$ErrorActionPreference = "Stop"

#region Helper Functions

function Invoke-AzCli {
    <#
    .SYNOPSIS
        Safely invokes an Azure CLI command and returns the output.
        Handles the PowerShell 5.1 issue where stderr from native commands
        becomes a terminating error when $ErrorActionPreference is 'Stop'.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $callerErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = & az @Arguments 2>&1
        $script:LastAzExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $callerErrorAction
    }

    return $output
}

function Invoke-AzRestWithRetry {
    <#
    .SYNOPSIS
        Calls az rest GET with retry logic for throttling (HTTP 429).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [int]$RetryCount = 3
    )

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        $response = Invoke-AzCli -Arguments @("rest", "--method", "GET", "--url", $Uri, "--only-show-errors")

        if ($script:LastAzExitCode -eq 0) {
            return ($response | ConvertFrom-Json)
        }

        $responseText = ($response | Out-String)
        if ($responseText -match "429" -or $responseText -match "throttl") {
            $waitSeconds = [math]::Pow(2, $attempt) * 5
            Write-Warning "Throttled on attempt $attempt for URI. Waiting $waitSeconds seconds..."
            Start-Sleep -Seconds $waitSeconds
        }
        else {
            return $null
        }
    }

    Write-Warning "Max retries ($RetryCount) exceeded for: $Uri"
    return $null
}

#endregion

#region Prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Cyan

# Ensure TLS 1.2 for PowerShell Gallery access
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Check Azure CLI is installed
$azVersionOutput = Invoke-AzCli -Arguments @("--version")
if ($script:LastAzExitCode -ne 0) {
    Write-Error "Azure CLI (az) is not installed or not in PATH. Install from https://aka.ms/installazurecli"
    exit 1
}

# Check Azure CLI is logged in
$null = Invoke-AzCli -Arguments @("account", "show", "--only-show-errors")
if ($script:LastAzExitCode -ne 0) {
    Write-Error "Azure CLI is not logged in. Run 'az login' first."
    exit 1
}

# Check/Install ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow

    # Ensure NuGet provider is available
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }

    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-Module -Name ImportExcel -Force -Scope CurrentUser -AllowClobber
}
Import-Module ImportExcel -ErrorAction Stop

Write-Host "Prerequisites OK." -ForegroundColor Green
#endregion

#region Get Subscriptions
Write-Host "`nRetrieving subscriptions..." -ForegroundColor Cyan

$subscriptionsJson = Invoke-AzCli -Arguments @("account", "list", "--all", "--output", "json", "--only-show-errors")
if ($script:LastAzExitCode -ne 0) {
    Write-Error "Failed to retrieve subscriptions. Ensure you are logged in with 'az login'."
    exit 1
}

$allSubscriptions = $subscriptionsJson | ConvertFrom-Json
$subscriptions = @($allSubscriptions | Where-Object {
    $_.state -eq "Enabled" -and $_.name -like $SubscriptionFilter
})

if ($subscriptions.Count -eq 0) {
    Write-Error "No enabled subscriptions found matching filter '$SubscriptionFilter'."
    exit 1
}

Write-Host "Found $($subscriptions.Count) enabled subscription(s)." -ForegroundColor Green

$subscriptionSummary = $subscriptions | ForEach-Object {
    [PSCustomObject]@{
        "Subscription Name" = $_.name
        "Subscription ID"   = $_.id
        "State"             = $_.state
        "Tenant ID"         = $_.tenantId
        "Is Default"        = $_.isDefault
    }
}
#endregion

#region Get Regions With Deployed Resources
Write-Host "`nDiscovering regions with deployed resources..." -ForegroundColor Cyan

$regionsPerSubscription = @{}

foreach ($sub in $subscriptions) {
    $null = Invoke-AzCli -Arguments @("account", "set", "--subscription", $sub.id, "--only-show-errors")

    $resourcesJson = Invoke-AzCli -Arguments @("resource", "list", "--query", "[].location", "--output", "json", "--only-show-errors")
    if ($script:LastAzExitCode -eq 0) {
        $deployedRegions = @(($resourcesJson | ConvertFrom-Json) | Sort-Object -Unique)
        $regionsPerSubscription[$sub.id] = $deployedRegions
    }
    else {
        Write-Warning "Could not list resources for subscription '$($sub.name)'. Skipping."
        $regionsPerSubscription[$sub.id] = @()
    }
}

$allRegions = @(($regionsPerSubscription.Values | ForEach-Object { $_ }) | Sort-Object -Unique)
Write-Host "Found $($allRegions.Count) region(s) with deployed resources across all subscriptions." -ForegroundColor Green
#endregion

#region Get Quotas
Write-Host "`nRetrieving quotas (this may take a while)..." -ForegroundColor Cyan

$quotaResults = New-Object System.Collections.Generic.List[PSCustomObject]
$totalIterations = ($regionsPerSubscription.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
$currentIteration = 0

foreach ($sub in $subscriptions) {
    $subId = $sub.id
    $subName = $sub.name
    $subRegions = $regionsPerSubscription[$subId]

    if ($subRegions.Count -eq 0) {
        Write-Host "  Skipping '$subName' - no deployed resources." -ForegroundColor DarkGray
        continue
    }

    foreach ($region in $subRegions) {
        $currentIteration++
        $percentComplete = [math]::Round(($currentIteration / $totalIterations) * 100, 1)
        Write-Progress -Activity "Retrieving quotas" `
            -Status "Subscription: $subName | Region: $region ($percentComplete%)" `
            -PercentComplete $percentComplete

        # Compute quotas (Microsoft.Compute)
        $computeUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Compute/locations/$region/usages?api-version=2023-03-01"
        $computeData = Invoke-AzRestWithRetry -Uri $computeUri -RetryCount $MaxRetries

        if ($computeData -and $computeData.value) {
            foreach ($item in $computeData.value) {
                $limit = $item.limit
                $currentUsage = $item.currentValue
                $pctUsed = if ($limit -gt 0) { [math]::Round(($currentUsage / $limit) * 100, 2) } else { 0 }

                $quotaResults.Add([PSCustomObject]@{
                    "Subscription Name" = $subName
                    "Subscription ID"   = $subId
                    "Region"            = $region
                    "Resource Provider"  = "Microsoft.Compute"
                    "Quota Name"        = $item.name.localizedValue
                    "Limit"             = $limit
                    "Current Usage"     = $currentUsage
                    "% Used"            = $pctUsed
                })
            }
        }

        # Network quotas (Microsoft.Network)
        $networkUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Network/locations/$region/usages?api-version=2023-05-01"
        $networkData = Invoke-AzRestWithRetry -Uri $networkUri -RetryCount $MaxRetries

        if ($networkData -and $networkData.value) {
            foreach ($item in $networkData.value) {
                $limit = $item.limit
                $currentUsage = $item.currentValue
                $pctUsed = if ($limit -gt 0) { [math]::Round(($currentUsage / $limit) * 100, 2) } else { 0 }

                $quotaResults.Add([PSCustomObject]@{
                    "Subscription Name" = $subName
                    "Subscription ID"   = $subId
                    "Region"            = $region
                    "Resource Provider"  = "Microsoft.Network"
                    "Quota Name"        = $item.name.localizedValue
                    "Limit"             = $limit
                    "Current Usage"     = $currentUsage
                    "% Used"            = $pctUsed
                })
            }
        }

        # Storage quotas (Microsoft.Storage)
        $storageUri = "https://management.azure.com/subscriptions/$subId/providers/Microsoft.Storage/locations/$region/usages?api-version=2023-01-01"
        $storageData = Invoke-AzRestWithRetry -Uri $storageUri -RetryCount $MaxRetries

        if ($storageData -and $storageData.value) {
            foreach ($item in $storageData.value) {
                $limit = $item.limit
                $currentUsage = $item.currentValue
                $pctUsed = if ($limit -gt 0) { [math]::Round(($currentUsage / $limit) * 100, 2) } else { 0 }

                $quotaResults.Add([PSCustomObject]@{
                    "Subscription Name" = $subName
                    "Subscription ID"   = $subId
                    "Region"            = $region
                    "Resource Provider"  = "Microsoft.Storage"
                    "Quota Name"        = $item.name.localizedValue
                    "Limit"             = $limit
                    "Current Usage"     = $currentUsage
                    "% Used"            = $pctUsed
                })
            }
        }
    }
}

Write-Progress -Activity "Retrieving quotas" -Completed
Write-Host "Collected $($quotaResults.Count) quota entries." -ForegroundColor Green
#endregion

#region Export to Excel
Write-Host "`nExporting to Excel: $OutputPath" -ForegroundColor Cyan

# Remove existing file if present
if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

# Sheet 1: Subscriptions
$subscriptionSummary | Export-Excel -Path $OutputPath `
    -WorksheetName "Subscriptions" `
    -AutoSize `
    -AutoFilter `
    -FreezeTopRow `
    -BoldTopRow `
    -TableStyle Medium6

# Sheet 2: Quotas
if ($quotaResults.Count -gt 0) {
    $quotaResults | Export-Excel -Path $OutputPath `
        -WorksheetName "Quotas" `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow `
        -TableStyle Medium6 `
        -Append

    # Add conditional formatting on % Used column
    $excel = Open-ExcelPackage -Path $OutputPath
    $ws = $excel.Workbook.Worksheets["Quotas"]
    $lastRow = $ws.Dimension.End.Row
    $pctColumn = 8  # Column H = "% Used"

    # Yellow for >= 70%
    Add-ConditionalFormatting -Worksheet $ws `
        -Range "H2:H$lastRow" `
        -RuleType GreaterThanOrEqual `
        -ConditionValue 70 `
        -BackgroundColor Yellow

    # Red for >= 90%
    Add-ConditionalFormatting -Worksheet $ws `
        -Range "H2:H$lastRow" `
        -RuleType GreaterThanOrEqual `
        -ConditionValue 90 `
        -BackgroundColor Red `
        -ForegroundColor White

    Close-ExcelPackage $excel
}
else {
    Write-Warning "No quota data collected. The Quotas sheet will be empty."
    @([PSCustomObject]@{ "Note" = "No quota data was retrieved. Check permissions and region availability." }) |
        Export-Excel -Path $OutputPath -WorksheetName "Quotas" -AutoSize -Append
}

Write-Host "`nDone! Report saved to: $OutputPath" -ForegroundColor Green
Write-Host "  - Subscriptions: $($subscriptions.Count)" -ForegroundColor White
Write-Host "  - Quota entries: $($quotaResults.Count)" -ForegroundColor White
Write-Host "  - Regions checked: $($allRegions.Count)" -ForegroundColor White
#endregion
