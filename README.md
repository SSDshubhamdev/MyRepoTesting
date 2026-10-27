# My1stCode — Excel to JSON with Azure Blob + GitHub Actions

This starter project accepts an Excel build sheet from a Windows laptop, uploads the original workbook to Azure Blob Storage, triggers GitHub Actions, converts the workbook to JSON, and uploads the JSON to a second container.

## Project paths

Copy these files into your existing project at `D:\VM build Automation\My1stCode`:

- `.github/workflows/vm-build-json.yml`
- `scripts/Build-VMJson.ps1`
- `scripts/Upload-And-Run-VMJson.ps1`

The parser is designed for the workbook structure reviewed for this project:
- `Build Sheet for Requestor`
- `Check List for Azure Team`

It reads the request details, VM rows, tags, access/domain fields, and checklist. If your customer changes the workbook layout, update and test the parser before production use.

## Azure setup (one-time)

1. Create a Storage Account with public blob access disabled.
2. Create private containers `vm-build-input` and `vm-build-json` (the local launcher also attempts to create them).
3. Give your signed-in Azure user the required blob data permissions to upload the Excel and create/verify containers.
4. Configure a Microsoft Entra app registration for GitHub OIDC. Its federated credential must match your actual GitHub repository and branch (`refs/heads/main`).
5. Give the GitHub OIDC app `Storage Blob Data Contributor` at the narrowest suitable scope (ideally the target storage account).
6. In GitHub repository Settings → Secrets and variables → Actions, add:
   - Secret `AZURE_CLIENT_ID`: OIDC app/client ID
   - Secret `AZURE_TENANT_ID`: tenant ID
   - Secret `AZURE_SUBSCRIPTION_ID`: subscription containing the Storage Account
   - Variable `STORAGE_ACCOUNT_NAME`: actual Storage Account name
7. Commit and push the workflow/scripts to branch `main`.

The GitHub workflow uses OIDC; do not add a client secret to GitHub for this workflow.

## Local laptop prerequisites (one-time)

Install and sign in to:
- Azure CLI: `az login`
- GitHub CLI: `gh auth login`
- Git, and clone your GitHub repo locally
- PowerShell 7 is recommended; Windows PowerShell can also run the launcher. The launcher installs the `ImportExcel` module for the local request-ID check.

Open PowerShell in your project folder:

```powershell
Set-Location 'D:\VM build Automation\My1stCode'
az login
gh auth login
```

Run the launcher; it opens a file picker if you don't supply `-ExcelPath`:

```powershell
.\scripts\Upload-And-Run-VMJson.ps1 `
  -RequestId 'RITM3856137' `
  -StorageAccountName 'YOUR_STORAGE_ACCOUNT_NAME' `
  -SubscriptionId 'YOUR_SUBSCRIPTION_ID' `
  -GitHubRepo 'OWNER/REPO'
```

Or supply the Excel path explicitly:

```powershell
.\scripts\Upload-And-Run-VMJson.ps1 `
  -ExcelPath 'D:\VM_BuildSheet.xlsx' `
  -RequestId 'RITM3856137' `
  -StorageAccountName 'YOUR_STORAGE_ACCOUNT_NAME' `
  -SubscriptionId 'YOUR_SUBSCRIPTION_ID' `
  -GitHubRepo 'OWNER/REPO'
```

The Request ID must exactly match `Service Request No.` in the Excel sheet. Each request writes to its own blob path:

- Input: `vm-build-input/<RequestId>/BuildSheet.xlsx`
- Output: `vm-build-json/<RequestId>/<RequestId>.json`

Uploads use overwrite disabled. For corrections, use a controlled revision process rather than overwriting the original request.

## Safety and current limitations

- This pipeline only generates and stores JSON. It does not deploy VMs or run Terraform.
- Approval is explicitly `NOT_VERIFIED`; JSON `deployment_ready` is `false`.
- Subscription display names are not automatically converted to subscription IDs.
- Image names are not automatically mapped to approved Marketplace image publisher/offer/SKU/version.
- The script performs basic required-field validation, not full customer-specific schema validation.
- Before production use, review the JSON schema with the customer/platform team, add authoritative approval integration, and test with non-production sample requests.
- The input workbook may contain personal/contact or access information. Keep containers private, limit access, and configure retention according to company policy.
