# Azure Public IP Inventory

Export a comprehensive inventory of all Azure public IPs across your subscription to Excel, including VNET associations, resource types, SKU details, and allocation methods. Designed for Azure DDoS Protection Network Protection planning and compliance auditing.

## Features

- **Complete Inventory** — Lists all public IPs (VMs, Firewalls, Bastions, NAT Gateways, App Gateways, etc.)
- **VNET Mapping** — Shows which VNET each IP is associated with (critical for DDoS planning)
- **Excel Export** — Formatted Excel table with auto-filter for easy analysis
- **Resource Details** — Includes SKU, allocation method, IP version, DNS labels, provisioning state
- **Multi-page** — Automatically handles subscriptions with >1000 public IPs

## Use Cases

- **DDoS Protection Planning** — Identify how many public IPs per VNET to plan Azure DDoS Protection Network Protection allocation
- **Compliance Auditing** — Track all public IPs and their associations for security/compliance reviews
- **Cost Analysis** — Identify unused or unattached public IPs
- **Migration Planning** — Inventory assets before cloud migration projects

## Requirements

- **PowerShell 7.0+** (Windows, Linux, macOS)
- **Azure Subscription** with read access to Network resources
- **Azure Modules:**
  - Az.Accounts
  - Az.ResourceGraph
  - ImportExcel

## Installation

```powershell
Install-Module Az.Accounts -Scope CurrentUser -Force
Install-Module Az.ResourceGraph -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```

## Usage

```powershell
Connect-AzAccount
Set-AzContext -Subscription "subscription-id"
.\Export-PublicIPs.ps1
```

Output: `PublicIPs_Inventory_YYYYMMDD_HHMM.xlsx`

## Output Columns

- ipAddress, resourceName, vnetName, subnetName
- associatedResource, associationType
- resourceGroup, location, sku, allocationMethod, ipVersion
- dnsLabel, provisioningState, tags

## DDoS Protection Planning

Each VNET protected by Azure DDoS Protection covers all public IPs in that VNET (typically 100 IP limit per plan).

## Troubleshooting

- **Not connected:** Run `Get-AzContext`
- **Module missing:** `Install-Module ImportExcel -Scope CurrentUser -Force`
- **Empty results:** Verify read permissions on Network resources

## Security

- Read-only script (no modifications to Azure resources)
- Uses Azure authentication via `Connect-AzAccount`
- Excel files contain sensitive data — store securely
- No external data transmission

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md)

## License

[MIT License](LICENSE)

## Code of Conduct

[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/)

---

**Created by:** [jsornsin](https://github.com/jsornsin)
