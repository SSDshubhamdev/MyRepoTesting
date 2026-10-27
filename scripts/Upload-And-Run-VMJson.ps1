
param(
    [Parameter(Mandatory = $true)]
    [string]$RequestId
)

$ErrorActionPreference = "Stop"

# Configuration
$StorageAccount = "shubhamstorage11"
$InputContainer = "shubhamc11"
$OutputContainer = "shubhamc111"
$ExcelBlobName = "VM_BuildSheet.xlsx"

# Paths
$ExcelPath = Join-Path $env:TEMP $ExcelBlobName
$OutputPath = Join-Path $env:TEMP "$RequestId.json"
$ConverterScript = Join-Path $PSScriptRoot "Build-VMJson.ps1"

try {
    Write-Host "=== STEP 1: Validate configuration ==="

    if (!(Get-Command az -ErrorAction SilentlyContinue)) {
        throw "Azure CLI is not installed."
    }

    if (!(Test-Path $ConverterScript)) {
        throw "Converter script not found: $ConverterScript"
    }

    Write-Host "=== STEP 2: Download Excel from Azure Blob ==="

    az storage blob download `
        --account-name $StorageAccount `
        --container-name $InputContainer `
        --name $ExcelBlobName `
        --file $ExcelPath `
        --auth-mode login `
        --only-show-errors `
        --output none

    if ($LASTEXITCODE -ne 0 -or !(Test-Path $ExcelPath)) {
        throw "Excel download failed."
    }

    Write-Host "Excel downloaded successfully."

    Write-Host "=== STEP 3: Convert Excel to JSON ==="

    & $ConverterScript `
        -ExcelPath $ExcelPath `
        -RequestId $RequestId `
        -OutputPath $OutputPath

    if ($LASTEXITCODE -ne 0 -or !(Test-Path $OutputPath)) {
        throw "Excel-to-JSON conversion failed."
    }

    Write-Host "JSON created successfully."

    Write-Host "=== STEP 4: Upload JSON to Azure Blob ==="

    $JsonBlobName = "$RequestId/$RequestId.json"

    az storage blob upload `
        --account-name $StorageAccount `
        --container-name $OutputContainer `
        --name $JsonBlobName `
        --file $OutputPath `
        --auth-mode login `
        --overwrite true `
        --only-show-errors `
        --output none

    if ($LASTEXITCODE -ne 0) {
        throw "JSON upload failed."
    }

    Write-Host "=== SUCCESS ==="
    Write-Host "Request ID: $RequestId"
    Write-Host "Output: $OutputContainer/$JsonBlobName"
}
catch {
    Write-Error $_
    exit 1
}
