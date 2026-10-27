
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [string]$RequestId,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

Import-Module ImportExcel

# --------------------------------------------------
# STEP 1: Validate the input file
# --------------------------------------------------

if (!(Test-Path -LiteralPath $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$workbook = Get-ExcelSheetInfo -Path $ExcelPath
$sheetNames = @($workbook | ForEach-Object { $_.Name })

$buildSheetName = "Build Sheet for Requestor"
$checklistSheetName = "Check List for Azure Team"

if ($buildSheetName -notin $sheetNames) {
    throw "Worksheet '$buildSheetName' was not found. Available sheets: $($sheetNames -join ', ')"
}

if ($checklistSheetName -notin $sheetNames) {
    throw "Worksheet '$checklistSheetName' was not found. Available sheets: $($sheetNames -join ', ')"
}

# Read worksheets without assuming headers start at row 1.
$buildRows = @(Import-Excel -Path $ExcelPath -WorksheetName $buildSheetName -NoHeader)
$checklistRows = @(Import-Excel -Path $ExcelPath -WorksheetName $checklistSheetName -NoHeader)

if ($buildRows.Count -eq 0) {
    throw "Build worksheet is empty."
}

# --------------------------------------------------
# STEP 2: Convert each worksheet row to cell values
# --------------------------------------------------

function Get-CellValues {
    param([object]$Row)

    $values = @(
        foreach ($property in $Row.PSObject.Properties) {
            $property.Value
        }
    )

    return ,$values
}

$rawBuildRows = [System.Collections.Generic.List[object]]::new()

foreach ($row in $buildRows) {
    $rawBuildRows.Add([object[]](Get-CellValues -Row $row))
}

$rawChecklistRows = [System.Collections.Generic.List[object]]::new()

foreach ($row in $checklistRows) {
    $rawChecklistRows.Add([object[]](Get-CellValues -Row $row))
}

# Normalise whitespace when comparing Excel labels.
function Normalize-Text {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return ([string]$Value -replace '\s+', ' ').Trim()
}

# --------------------------------------------------
# STEP 3: Read label/value fields
# --------------------------------------------------

function Get-LabelValue {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Label
    )

    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $row = $Rows[$r]

        for ($c = 0; $c -lt $row.Count; $c++) {
            if ((Normalize-Text $row[$c]) -eq (Normalize-Text $Label)) {
                for ($next = $c + 1; $next -lt $row.Count; $next++) {
                    $value = Normalize-Text $row[$next]

                    if ($value -ne "") {
                        return $value
                    }
                }
            }
        }
    }

    return $null
}

$excelRequestId = Get-LabelValue -Rows $rawBuildRows -Label "Service Request No."

if ([string]::IsNullOrWhiteSpace($excelRequestId)) {
    throw "Service Request No. was not found in the Excel file."
}

if ($excelRequestId -ne $RequestId.Trim()) {
    throw "Request ID mismatch. Excel='$excelRequestId'; workflow='$RequestId'."
}

# --------------------------------------------------
# STEP 4: Locate the VM table header
# --------------------------------------------------

$headerRowIndex = -1
$columnMap = @{}

for ($r = 0; $r -lt $rawBuildRows.Count; $r++) {
    $candidateMap = @{}
    $row = $rawBuildRows[$r]

    for ($c = 0; $c -lt $row.Count; $c++) {
        $header = Normalize-Text $row[$c]

        if ($header -ne "") {
            $candidateMap[$header] = $c
        }
    }

    if (
        $candidateMap.ContainsKey("VM Name") -and
        $candidateMap.ContainsKey("Subscription") -and
        $candidateMap.ContainsKey("VM Size")
    ) {
        $headerRowIndex = $r
        $columnMap = $candidateMap
        break
    }
}

if ($headerRowIndex -lt 0) {
    throw "VM table headers were not found. Check the worksheet's column names."
}

function Get-ColumnValue {
    param(
        [object[]]$Row,
        [hashtable]$Columns,
        [string]$ColumnName
    )

    if (!$Columns.ContainsKey($ColumnName)) {
        return $null
    }

    $index = [int]$Columns[$ColumnName]

    if ($index -ge $Row.Count) {
        return $null
    }

    $value = Normalize-Text $Row[$index]

    if ($value -eq "" -or $value -eq "-") {
        return $null
    }

    return $value
}

# --------------------------------------------------
# STEP 5: Extract VM records
# --------------------------------------------------

$vms = [System.Collections.Generic.List[object]]::new()

for ($r = $headerRowIndex + 1; $r -lt $rawBuildRows.Count; $r++) {
    $row = $rawBuildRows[$r]

    $vmName = Get-ColumnValue -Row $row -Columns $columnMap -ColumnName "VM Name"

    if ([string]::IsNullOrWhiteSpace($vmName)) {
        continue
    }

    $vms.Add([ordered]@{
        request_id                 = $RequestId
        vm_name                    = $vmName
        os_image_display_name      = Get-ColumnValue $row $columnMap "Operating System / Image Name"
        subscription_name          = Get-ColumnValue $row $columnMap "Subscription"
        region                     = Get-ColumnValue $row $columnMap "Region"
        environment                = Get-ColumnValue $row $columnMap "Environment"
        business_unit              = Get-ColumnValue $row $columnMap "Bussiness Unit"
        application_name            = Get-ColumnValue $row $columnMap "Application Name"
        resource_group_name         = Get-ColumnValue $row $columnMap "Resource Group Name"
        vm_size_display_name        = Get-ColumnValue $row $columnMap "VM Size"
        vm_size                    = $null
        vnet_name                  = Get-ColumnValue $row $columnMap "Vnet"
        subnet_name                = Get-ColumnValue $row $columnMap "SubNet"
        nsg_name                   = Get-ColumnValue $row $columnMap "NSG"
        private_static_ip          = Get-ColumnValue $row $columnMap "Private Static IP"
        diagnostic_storage_account = Get-ColumnValue $row $columnMap "Diagnostic Storage Account"
        additional_disk_count      = Get-ColumnValue $row $columnMap "No. of Additional Disk"
        additional_disk_type       = Get-ColumnValue $row $columnMap "Additional Disk Type"
        additional_disk_size       = Get-ColumnValue $row $columnMap "Size Additional Disk"
        managed_disk               = Get-ColumnValue $row $columnMap "Managed Disk"
        backup_vault_name          = Get-ColumnValue $row $columnMap "Backup Vault Name"
        backup_policy_name         = Get-ColumnValue $row $columnMap "Backup Policy Name"
        wipro_support              = Get-ColumnValue $row $columnMap "Wipro Support"
        availability_set           = Get-ColumnValue $row $columnMap "Availability Set"
        availability_set_name      = Get-ColumnValue $row $columnMap "Availability Set Name"
        asr                        = Get-ColumnValue $row $columnMap "ASR"
    })
}

if ($vms.Count -eq 0) {
    throw "No VM records were found in the VM table."
}

# --------------------------------------------------
# STEP 6: Read request and tag details
# --------------------------------------------------

$requestDetails = [ordered]@{
    request_id                    = $RequestId
    requestor_name                = Get-LabelValue $rawBuildRows "Build Requestor Name"
    requestor_email               = Get-LabelValue $rawBuildRows "Build Requestor Email ID"
    application_owner_name        = Get-LabelValue $rawBuildRows "Application / Server Owner Name"
    application_owner_email       = Get-LabelValue $rawBuildRows "Application Owner Email ID"
    application_team_dl           = Get-LabelValue $rawBuildRows "Application Team DL"
    application_type              = Get-LabelValue $rawBuildRows "Application Type"
    build_date                    = Get-LabelValue $rawBuildRows "Build Date"
    tag_environment               = Get-LabelValue $rawBuildRows "Environment"
    metal_band                    = Get-LabelValue $rawBuildRows "Metal Band"
    os_type                       = Get-LabelValue $rawBuildRows "OS Type"
    project_name                  = Get-LabelValue $rawBuildRows "Project Name"
    rtb_status                    = Get-LabelValue $rawBuildRows "RTB status"
    wipro_managed                 = Get-LabelValue $rawBuildRows "Wipro_managed"
    wipro_managed_env             = Get-LabelValue $rawBuildRows "Wipro_Managed_env"
    wipro_managedservices_tier    = Get-LabelValue $rawBuildRows "Wipro_Managedservices_tier"
    live_ou                       = Get-LabelValue $rawBuildRows "Live OU:"
    staging_ou                    = Get-LabelValue $rawBuildRows "Staging OU:"
    application_admin_group       = Get-LabelValue $rawBuildRows "Application Admin Group:"
}

# --------------------------------------------------
# STEP 7: Keep checklist information for review
# --------------------------------------------------

$checklist = @(
    foreach ($row in $rawChecklistRows) {
        $cells = @(
            foreach ($cell in $row) {
                $value = Normalize-Text $cell
                if ($value -ne "") { $value }
            }
        )

        if ($cells.Count -gt 0) {
            [string]::Join(" | ", $cells)
        }
    }
)

# --------------------------------------------------
# STEP 8: Build validation notes
# --------------------------------------------------

$validationNotes = [System.Collections.Generic.List[object]]::new()

$validationNotes.Add([ordered]@{
    field  = "service_request_approval"
    status = "NOT_VERIFIED"
    action = "Confirm approval in the approved service-management system before deployment."
})

$validationNotes.Add([ordered]@{
    field  = "azure_subscription_id"
    status = "NEEDS_RESOLUTION"
    action = "The workbook provides a subscription name; resolve and verify its Azure subscription ID."
})

$validationNotes.Add([ordered]@{
    field  = "os_image"
    status = "NEEDS_MAPPING"
    action = "Confirm the approved Azure image publisher, offer, SKU, and version."
})

$validationNotes.Add([ordered]@{
    field  = "network_and_security"
    status = "NEEDS_VALIDATION"
    action = "Resolve and verify VNet, subnet, NSG, IP, and resource IDs before deployment."
})

$validationNotes.Add([ordered]@{
    field  = "disks_backup_and_asr"
    status = "NEEDS_VALIDATION"
    action = "Confirm disk configuration, backup policy, vault, and ASR requirements."
})

$validationNotes.Add([ordered]@{
    field  = "mandatory_tags_and_access"
    status = "NEEDS_VALIDATION"
    action = "Review blank tags, access requirements, domain/OU settings, and required groups."
})

# --------------------------------------------------
# STEP 9: Create JSON
# --------------------------------------------------

$result = [ordered]@{
    request              = $requestDetails
    vms                  = @($vms.ToArray())
    checklist_for_review = $checklist
    validation_notes     = @($validationNotes.ToArray())
    approval = [ordered]@{
        status = "NOT_VERIFIED"
    }
    deployment_ready = $false
    deployment_status = "DRAFT_NOT_READY_FOR_TERRAFORM_APPLY"
    important_note = "This JSON is an extracted draft. It is not approval verification or deployment authorization. Validate required Azure IDs, image, network, disks, backup, access, and tags before any Terraform apply."
}

$outputDirectory = Split-Path -Parent $OutputPath

if (![string]::IsNullOrWhiteSpace($outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$json = ConvertTo-Json -InputObject $result -Depth 20
Set-Content -LiteralPath $OutputPath -Value $json -Encoding utf8

Write-Host "Excel validation completed."
Write-Host "Request ID: $RequestId"
Write-Host "VM records extracted: $($vms.Count)"
Write-Host "Approval status: NOT_VERIFIED"
Write-Host "Deployment ready: false"
Write-Host "JSON output: $OutputPath"
