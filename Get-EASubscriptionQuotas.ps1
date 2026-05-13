<#
.SYNOPSIS
    Lists all EA subscriptions and their resource quotas, exporting results to an Excel file.

.DESCRIPTION
    This script enumerates all Azure subscriptions accessible to the current user,
    retrieves quota usage for all resource providers across all regions, and exports
    the data to a formatted Excel workbook.

.PARAMETER OutputPath
    Path for the output Excel file. Defaults to EA-Subscription-Quotas-<date>.xlsx in the current directory.

.PARAMETER SubscriptionFilter
    Optional. Filter subscriptions by name pattern (supports wildcards).

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

#region Prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Cyan

# Check Azure CLI
try {
    $null = az account show 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Azure CLI is not logged in. Run 'az login' first."
        exit 1
    }
}
catch {
    Write-Error "Azure CLI (az) is not installed or not in PATH. Install from https://aka.ms/installazurecli"
    exit 1
}

# Check/Install ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser -AllowClobber
}
Import-Module ImportExcel

Write-Host "Prerequisites OK." -ForegroundColor Green
#endregion

#region Helper Functions
function Invoke-AzRestWithRetry {
    param(
        [string]$Uri,
        [int]$MaxRetries = 3
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $response = az rest --method GET --url $Uri 2>&1
        if ($LASTEXITCODE -eq 0) {
            return ($response | ConvertFrom-Json)
        }

        $responseText = $response -join " "
        if ($responseText -match "429" -or $responseText -match "throttl") {
            $waitSeconds = [math]::Pow(2, $attempt) * 5
            Write-Warning "Throttled on attempt $attempt. Waiting $waitSeconds seconds..."
            Start-Sleep -Seconds $waitSeconds
        }
        else {
            return $null
        }
    }
    Write-Warning "Max retries exceeded for: $Uri"
    return $null
}
#endregion

#region Get Subscriptions
Write-Host "`nRetrieving subscriptions..." -ForegroundColor Cyan

$subscriptionsJson = az account list --all --output json 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to retrieve subscriptions: $subscriptionsJson"
    exit 1
}

$allSubscriptions = $subscriptionsJson | ConvertFrom-Json
$subscriptions = $allSubscriptions | Where-Object {
    $_.state -eq "Enabled" -and $_.name -like $SubscriptionFilter
}

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
    az account set --subscription $sub.id 2>&1 | Out-Null

    $resourcesJson = az resource list --query "[].location" --output json 2>&1
    if ($LASTEXITCODE -eq 0) {
        $deployedRegions = ($resourcesJson | ConvertFrom-Json) | Sort-Object -Unique
        $regionsPerSubscription[$sub.id] = $deployedRegions
    }
    else {
        Write-Warning "Could not list resources for subscription '$($sub.name)'. Skipping."
        $regionsPerSubscription[$sub.id] = @()
    }
}

$allRegions = ($regionsPerSubscription.Values | ForEach-Object { $_ }) | Sort-Object -Unique
Write-Host "Found $($allRegions.Count) region(s) with deployed resources across all subscriptions." -ForegroundColor Green
#endregion

#region Get Quotas
Write-Host "`nRetrieving quotas (this may take a while)..." -ForegroundColor Cyan

$quotaResults = [System.Collections.Generic.List[PSCustomObject]]::new()
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
        $computeData = Invoke-AzRestWithRetry -Uri $computeUri -MaxRetries $MaxRetries

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
        $networkData = Invoke-AzRestWithRetry -Uri $networkUri -MaxRetries $MaxRetries

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
        $storageData = Invoke-AzRestWithRetry -Uri $storageUri -MaxRetries $MaxRetries

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
