[CmdletBinding()]
param(
    [string]$ExcelPath,
    [string]$RequestId,
    [string]$StorageAccountName,
    [string]$SubscriptionId,
    [string]$GitHubRepo
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

# Windows file picker if ExcelPath is not supplied.
if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select the VM build sheet Excel file'
    $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'No Excel file selected. Cancelled.'
    }
    $ExcelPath = $dialog.FileName
}
if (!(Test-Path -LiteralPath $ExcelPath -PathType Leaf)) { throw "Excel file not found: $ExcelPath" }
if ([IO.Path]::GetExtension($ExcelPath) -ne '.xlsx') { throw 'Select an .xlsx file.' }

if ([string]::IsNullOrWhiteSpace($RequestId)) { $RequestId = Read-Host 'Enter Request ID exactly as in Excel (e.g. RITM3856137)' }
if ($RequestId -notmatch '^[A-Za-z0-9._-]+$') { throw 'Request ID can contain only letters, numbers, dot, underscore and hyphen.' }
if ([string]::IsNullOrWhiteSpace($StorageAccountName)) { $StorageAccountName = Read-Host 'Enter Azure Storage Account name' }
if ([string]::IsNullOrWhiteSpace($SubscriptionId)) { $SubscriptionId = Read-Host 'Enter Azure Subscription ID containing the Storage Account' }
if ([string]::IsNullOrWhiteSpace($GitHubRepo)) {
    $remote = (git remote get-url origin 2>$null | Select-Object -First 1)
    if ($LASTEXITCODE -eq 0 -and $remote) {
        if ($remote -match 'github\.com[:/]([^/]+/[^/.]+?)(?:\.git)?$') { $GitHubRepo = $Matches[1] }
    }
}
if ([string]::IsNullOrWhiteSpace($GitHubRepo)) { $GitHubRepo = Read-Host 'Enter GitHub repository as OWNER/REPO (e.g. my-org/My1stCode)' }

# Read the request ID from the Excel before upload, so a mismatched ID is caught early.
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
}
Import-Module ImportExcel -ErrorAction Stop
$package = $null
try {
    $package = Open-ExcelPackage -Path $ExcelPath
    $sheet = $package.Workbook.Worksheets['Build Sheet for Requestor']
    if ($null -eq $sheet) { throw "Worksheet 'Build Sheet for Requestor' not found." }
    $excelRequestId = ''
    for ($r = 1; $r -le $sheet.Dimension.End.Row; $r++) {
        $label = ([string]$sheet.Cells[$r,3].Text).Trim().TrimEnd(':')
        if ($label -ieq 'Service Request No.') { $excelRequestId = ([string]$sheet.Cells[$r,4].Text).Trim(); break }
    }
    if ($excelRequestId -ne $RequestId) {
        throw "Request ID mismatch. Entered '$RequestId', Excel contains '$excelRequestId'."
    }
}
finally {
    if ($null -ne $package) { $package.Dispose() }
}

foreach ($tool in @('az','gh','git')) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) { throw "$tool CLI is not installed or not available in PATH." }
}

# Sign in interactively if needed; the local user needs blob data upload rights.
az account show --output none 2>$null
if ($LASTEXITCODE -ne 0) {
    az login
    if ($LASTEXITCODE -ne 0) { throw 'Azure login failed.' }
}
az account set --subscription $SubscriptionId
if ($LASTEXITCODE -ne 0) { throw 'Could not select Azure subscription. Check the subscription ID and your access.' }

# Ensure containers exist. Requires Storage Blob Data Contributor (or equivalent) for the signed-in user.
foreach ($container in @('vm-build-input','vm-build-json')) {
    az storage container create --auth-mode login --account-name $StorageAccountName --name $container --only-show-errors --output none
    if ($LASTEXITCODE -ne 0) { throw "Could not create/verify container '$container'. Check permissions and account name." }
}

# Upload original file without overwriting an existing request's source.
$blobName = "$RequestId/BuildSheet.xlsx"
Write-Host "Uploading Excel to $StorageAccountName/vm-build-input/$blobName ..."
az storage blob upload `
    --auth-mode login `
    --account-name $StorageAccountName `
    --container-name 'vm-build-input' `
    --name $blobName `
    --file $ExcelPath `
    --overwrite false `
    --only-show-errors
if ($LASTEXITCODE -ne 0) { throw 'Excel upload failed. If this request already exists, do not overwrite it; use a new request ID or an approved revision process.' }

# Check GitHub CLI authentication before starting the workflow.
gh auth status
if ($LASTEXITCODE -ne 0) { throw "GitHub CLI is not authenticated. Run 'gh auth login' and retry. Excel has already uploaded." }

Write-Host "Triggering workflow in $GitHubRepo ..."
gh workflow run vm-build-json.yml --repo $GitHubRepo --ref main -f "request_id=$RequestId"
if ($LASTEXITCODE -ne 0) { throw 'Workflow trigger failed. Excel is uploaded; check GitHub authentication/repository/workflow before retrying.' }

Write-Host ''
Write-Host 'SUCCESS: Excel uploaded and GitHub Actions workflow triggered.' -ForegroundColor Green
Write-Host "Request ID: $RequestId"
Write-Host "Input: vm-build-input/$blobName"
Write-Host "Expected output: vm-build-json/$RequestId/$RequestId.json"
Write-Host "View run: https://github.com/$GitHubRepo/actions"
