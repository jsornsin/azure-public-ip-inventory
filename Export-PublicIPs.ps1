# Export Azure Public IP Inventory for DDoS Protection Planning
# This script queries Azure Resource Graph to inventory all public IPs across your subscription
# and exports them to Excel with VNET associations for DDoS Network Protection planning.
#
# Requirements:
#   - PowerShell 7.0+ (cross-platform compatible)
#   - Az.Accounts module (Install-Module Az.Accounts -Scope CurrentUser)
#   - Az.ResourceGraph module (Install-Module Az.ResourceGraph -Scope CurrentUser)
#   - ImportExcel module (Install-Module ImportExcel -Scope CurrentUser)
#
# Usage:
#   1. Connect-AzAccount -TenantId <your-tenant-id>
#   2. Set-AzContext -Subscription <your-subscription-id>
#   3. .\Export-PublicIPs.ps1
#
# Output:
#   Excel file with columns: IP Address, VNET, Subnet, Resource Group, Associated Resource, Type, SKU, etc.
#   Formatted as a table with auto-filter for easy analysis.

# ============================================================================
# Configuration
# ============================================================================

# Get current context — uses the subscription you set with Set-AzContext
$context = Get-AzContext
if (-not $context) {
    Write-Host "ERROR: Not connected to Azure. Run Connect-AzAccount first." -ForegroundColor Red
    exit 1
}

$subscriptionId = $context.Subscription.Id
$subscriptionName = $context.Subscription.Name

$outputFile = "PublicIPs_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"

Write-Host "Connected to: $subscriptionName ($subscriptionId)" -ForegroundColor Green

# ============================================================================
# Module Checks
# ============================================================================

if (-not (Get-Module ImportExcel -ListAvailable)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel -ErrorAction Stop

# ============================================================================
# Helper Function: Query Azure Resource Graph
# ============================================================================

function Invoke-ArgQuery {
    param(
        [string]$Query,
        [string]$SubscriptionId
    )
    
    $all = @()
    $skipToken = $null
    $pageCount = 0
    
    do {
        $pageCount++
        $params = @{
            Query = $Query
            Subscription = $SubscriptionId
            First = 1000
        }
        if ($skipToken) { $params['SkipToken'] = $skipToken }
        
        $response = Search-AzGraph @params
        $all += $response
        $skipToken = $response.SkipToken
        
        Write-Host "  Page $pageCount: $($response.Count) records" -ForegroundColor DarkGray
    } while ($skipToken)
    
    return $all
}

# ============================================================================
# Queries
# ============================================================================

Write-Host "`nQuerying Azure Resource Graph..." -ForegroundColor Cyan

Write-Host "  1. Public IPs..." -ForegroundColor White
$pips = Invoke-ArgQuery `
    "resources | where type =~ 'microsoft.network/publicipaddresses' | project pipId = tolower(id), resourceName = name, resourceGroup, subscriptionId, location, ipAddress = tostring(properties.ipAddress), sku = tostring(sku.name), allocationMethod = tostring(properties.publicIPAllocationMethod), ipVersion = tostring(properties.publicIPAddressVersion), dnsLabel = tostring(properties.dnsSettings.fqdn), provisioningState = tostring(properties.provisioningState), tags" `
    $subscriptionId

Write-Host "  2. NIC associations (VMs)..." -ForegroundColor White
$nics = Invoke-ArgQuery `
    "resources | where type =~ 'microsoft.network/networkinterfaces' | mv-expand ipConfig = properties.ipConfigurations | extend pipId = tolower(tostring(ipConfig.properties.publicIPAddress.id)) | extend subnetId = tostring(ipConfig.properties.subnet.id) | where isnotempty(pipId) | parse subnetId with * '/virtualNetworks/' vnetName '/subnets/' subnetName | project pipId, associatedResource = name, associationType = 'NIC (VM)', vnetName, subnetName" `
    $subscriptionId

Write-Host "  3. Azure Firewall associations..." -ForegroundColor White
$fws = Invoke-ArgQuery `
    "resources | where type =~ 'microsoft.network/azurefirewalls' | mv-expand ipConfig = properties.ipConfigurations | extend pipId = tolower(tostring(ipConfig.properties.publicIPAddress.id)) | extend subnetId = tostring(ipConfig.properties.subnet.id) | where isnotempty(pipId) | parse subnetId with * '/virtualNetworks/' vnetName '/subnets/' subnetName | project pipId, associatedResource = name, associationType = 'Azure Firewall', vnetName, subnetName" `
    $subscriptionId

Write-Host "  4. Bastion associations..." -ForegroundColor White
$bastions = Invoke-ArgQuery `
    "resources | where type =~ 'microsoft.network/bastionhosts' | mv-expand ipConfig = properties.ipConfigurations | extend pipId = tolower(tostring(ipConfig.properties.publicIPAddress.id)) | extend subnetId = tostring(ipConfig.properties.subnet.id) | where isnotempty(pipId) | parse subnetId with * '/virtualNetworks/' vnetName '/subnets/' subnetName | project pipId, associatedResource = name, associationType = 'Bastion', vnetName, subnetName" `
    $subscriptionId

Write-Host "  5. NAT Gateway associations..." -ForegroundColor White
$natgws = Invoke-ArgQuery `
    "resources | where type =~ 'microsoft.network/natgateways' | mv-expand pip = properties.publicIpAddresses | mv-expand sn = properties.subnets | extend pipId = tolower(tostring(pip.id)) | extend subnetId = tostring(sn.id) | where isnotempty(pipId) | parse subnetId with * '/virtualNetworks/' vnetName '/subnets/' subnetName | project pipId, associatedResource = name, associationType = 'NAT Gateway', vnetName, subnetName" `
    $subscriptionId

Write-Host "  6. Application Gateway associations..." -ForegroundColor White
$agws = Invoke-ArgQuery `
    "resources | where type =~ 'microsoft.network/applicationgateways' | mv-expand fe = properties.frontendIPConfigurations | mv-expand gw = properties.gatewayIPConfigurations | extend pipId = tolower(tostring(fe.properties.publicIPAddress.id)) | extend subnetId = tostring(gw.properties.subnet.id) | where isnotempty(pipId) | parse subnetId with * '/virtualNetworks/' vnetName '/subnets/' subnetName | project pipId, associatedResource = name, associationType = 'Application Gateway', vnetName, subnetName" `
    $subscriptionId

Write-Host "  7. Subscription details..." -ForegroundColor White
$subs = Invoke-ArgQuery `
    "resourcecontainers | where type =~ 'microsoft.resources/subscriptions' | project subscriptionId, subscriptionName = name" `
    $subscriptionId

# ============================================================================
# Join Data in PowerShell
# ============================================================================

Write-Host "`nProcessing results..." -ForegroundColor Cyan

$assocLookup = @{}
foreach ($row in @($nics) + @($fws) + @($bastions) + @($natgws) + @($agws)) {
    if ($row.pipId) { $assocLookup[$row.pipId] = $row }
}

$subLookup = @{}
foreach ($row in $subs) { $subLookup[$row.subscriptionId] = $row.subscriptionName }

$results = foreach ($pip in $pips) {
    $assoc = $assocLookup[$pip.pipId]
    [PSCustomObject]@{
        subscriptionId     = $pip.subscriptionId
        subscriptionName   = $subLookup[$pip.subscriptionId]
        resourceGroup      = $pip.resourceGroup
        location           = $pip.location
        ipAddress          = $pip.ipAddress
        resourceName       = $pip.resourceName
        sku                = $pip.sku
        allocationMethod   = $pip.allocationMethod
        ipVersion          = $pip.ipVersion
        dnsLabel           = $pip.dnsLabel
        provisioningState  = $pip.provisioningState
        associatedResource = $assoc.associatedResource
        associationType    = $assoc.associationType
        vnetName           = $assoc.vnetName
        subnetName         = $assoc.subnetName
        tags               = ($pip.tags | ConvertTo-Json -Compress)
    }
}

# ============================================================================
# Export to Excel
# ============================================================================

Write-Host "`nExporting to Excel..." -ForegroundColor Cyan

$results | Sort-Object subscriptionName, vnetName, resourceGroup, resourceName |
    Export-Excel -Path $outputFile -WorksheetName "Public IPs" -TableName "PublicIPs" `
        -TableStyle "Medium2" -AutoFilter

Write-Host "✓ Exported $($results.Count) records to '$outputFile'" -ForegroundColor Green
Write-Host "  Location: $(Resolve-Path $outputFile)" -ForegroundColor Gray
