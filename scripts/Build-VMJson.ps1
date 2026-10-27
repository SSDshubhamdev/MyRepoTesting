[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ExcelPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string]$RequestId
)

$ErrorActionPreference = 'Stop'

function Get-CellText {
    param($Worksheet, [int]$Row, [int]$Column)
    $value = $Worksheet.Cells[$Row, $Column].Text
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim()
}

function Get-ValueBesideLabel {
    param($Worksheet, [string]$Label, [int]$LabelColumn = 3, [int]$ValueColumn = 4)
    for ($row = 1; $row -le $Worksheet.Dimension.End.Row; $row++) {
        $candidate = (Get-CellText -Worksheet $Worksheet -Row $row -Column $LabelColumn).Trim().TrimEnd(':')
        if ($candidate -and $candidate.Equals($Label.Trim().TrimEnd(':'), [StringComparison]::OrdinalIgnoreCase)) {
            return (Get-CellText -Worksheet $Worksheet -Row $row -Column $ValueColumn)
        }
    }
    return ''
}

function Convert-ToTagKey {
    param([string]$Name)
    switch -Regex ($Name.Trim()) {
        '^Application Type$' { return 'ApplicationType' }
        '^Environment$' { return 'Environment' }
        '^Metal Band$' { return 'MetalBand' }
        '^OS Type$' { return 'OSType' }
        '^Project Name$' { return 'ProjectName' }
        '^Build Date$' { return 'BuildDate' }
        '^RTB status$' { return 'RTBStatus' }
        '^Wipro_managed$' { return 'Wipro_managed' }
        '^Wipro_Managed_env$' { return 'Wipro_Managed_env' }
        '^Wipro_Managedservices_tier$' { return 'Wipro_Managedservices_tier' }
        default { return ($Name -replace '[^A-Za-z0-9_]', '') }
    }
}

$package = $null
try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "PowerShell module 'ImportExcel' is missing. Install it with: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel -ErrorAction Stop

    $package = Open-ExcelPackage -Path $ExcelPath
    $requestSheet = $package.Workbook.Worksheets['Build Sheet for Requestor']
    $checklistSheet = $package.Workbook.Worksheets['Check List for Azure Team']

    if ($null -eq $requestSheet) { throw "Worksheet 'Build Sheet for Requestor' was not found." }
    if ($null -eq $checklistSheet) { throw "Worksheet 'Check List for Azure Team' was not found." }

    $sheetRequestId = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Service Request No.'
    if ([string]::IsNullOrWhiteSpace($sheetRequestId)) { throw 'Service Request No. is blank in the Excel sheet.' }
    if ($sheetRequestId -ne $RequestId) {
        throw "Request ID mismatch. Workflow RequestId='$RequestId', Excel Service Request No.='$sheetRequestId'. Use the exact ID from the sheet."
    }

    $requestorName = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Build Requestor Name'
    $requestorEmail = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Build Requestor Email ID'
    $ownerName = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Application / Server Owner Name'
    $ownerEmail = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Application Owner Email ID'
    $teamDl = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Application Team DL'

    # Read VM rows under the known header at row 11. Blank VM name ends the table.
    $vms = @()
    for ($row = 12; $row -le $requestSheet.Dimension.End.Row; $row++) {
        $vmName = Get-CellText -Worksheet $requestSheet -Row $row -Column 6
        $osName = Get-CellText -Worksheet $requestSheet -Row $row -Column 2
        if ([string]::IsNullOrWhiteSpace($vmName)) {
            if ($row -gt 12) { break }
            continue
        }
        $vms += [ordered]@{
            request_id          = $RequestId
            row_number          = $row
            os_image_display_name = $osName
            subscription_name   = Get-CellText -Worksheet $requestSheet -Row $row -Column 3
            region_display_name = Get-CellText -Worksheet $requestSheet -Row $row -Column 4
            resource_group      = Get-CellText -Worksheet $requestSheet -Row $row -Column 5
            vm_name             = $vmName
            vm_size_display_name = Get-CellText -Worksheet $requestSheet -Row $row -Column 7
            vnet_name           = Get-CellText -Worksheet $requestSheet -Row $row -Column 8
            subnet_name         = Get-CellText -Worksheet $requestSheet -Row $row -Column 9
        }
    }
    if ($vms.Count -eq 0) { throw 'No VM rows were found. Check the VM table (expected headers on row 11 and data from row 12).' }

    # Read tag key/value pairs from columns C:D, rows 17-26.
    $tags = [ordered]@{}
    for ($row = 17; $row -le 26; $row++) {
        $tagName = Get-CellText -Worksheet $requestSheet -Row $row -Column 3
        $tagValue = Get-CellText -Worksheet $requestSheet -Row $row -Column 4
        if (-not [string]::IsNullOrWhiteSpace($tagName)) {
            $tags[(Convert-ToTagKey -Name $tagName)] = if ($tagValue) { $tagValue } else { $null }
        }
    }

    $access = [ordered]@{
        local_id_service_id_or_group = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'LOCAL ID / Service ID / Local Group Access'
        local_id_name_to_create      = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Local ID Name if needs to Create'
        local_id_access_level        = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Level of Access for Local ID'
        domain_id_or_group_access    = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Domain ID / Group Access'
        nmpc_group_access            = Get-ValueBesideLabel -Worksheet $requestSheet -Label "NMPC ID's / Group Needs access"
        nmpc_group_access_level      = Get-ValueBesideLabel -Worksheet $requestSheet -Label "Level of Access for  NMPC ID's / Group"
        sql_access_level             = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'SQL Level off Access'
        nmpc_sql_group_access        = Get-ValueBesideLabel -Worksheet $requestSheet -Label "NMPC ID's / Group Needs SQL access"
        nmpc_sql_access_level        = Get-ValueBesideLabel -Worksheet $requestSheet -Label "Level of Access for  NMPC ID's / Group  SQL access"
    }

    $domain = [ordered]@{
        live_ou                = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Live OU'
        staging_ou             = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Staging OU'
        application_admin_group = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Application Admin Group'
        domain                 = Get-ValueBesideLabel -Worksheet $requestSheet -Label 'Domain'
    }

    $checklist = @()
    for ($row = 4; $row -le $checklistSheet.Dimension.End.Row; $row++) {
        $action = Get-CellText -Worksheet $checklistSheet -Row $row -Column 2
        if ([string]::IsNullOrWhiteSpace($action)) { continue }
        $checklist += [ordered]@{
            action        = $action
            status        = (Get-CellText -Worksheet $checklistSheet -Row $row -Column 3) -replace '^$', $null
            responsibility = (Get-CellText -Worksheet $checklistSheet -Row $row -Column 4) -replace '^$', $null
            performed_by  = (Get-CellText -Worksheet $checklistSheet -Row $row -Column 5) -replace '^$', $null
        }
    }

    $validationErrors = @()
    if ([string]::IsNullOrWhiteSpace($requestorName)) { $validationErrors += 'Build Requestor Name is blank.' }
    if ([string]::IsNullOrWhiteSpace($requestorEmail)) { $validationErrors += 'Build Requestor Email ID is blank.' }
    foreach ($vm in $vms) {
        foreach ($field in @('vm_name','os_image_display_name','subscription_name','region_display_name','resource_group','vm_size_display_name','vnet_name','subnet_name')) {
            if ([string]::IsNullOrWhiteSpace([string]$vm[$field])) {
                $validationErrors += "VM row $($vm.row_number): required field '$field' is blank."
            }
        }
    }
    if ($validationErrors.Count -gt 0) {
        throw ("Validation failed:`n - " + ($validationErrors -join "`n - "))
    }

    # Do not claim customer approval or map display names to real Azure resource IDs.
    $validationNotes = @(
        [ordered]@{ field = 'approval'; status = 'NOT_VERIFIED'; action = 'Confirm approval status in the approved request system before deployment.' },
        [ordered]@{ field = 'subscription_id'; status = 'NEEDS_MAPPING'; action = 'Map the subscription display name to an approved Azure subscription ID.' },
        [ordered]@{ field = 'os_image'; status = 'NEEDS_MAPPING'; action = 'Confirm the approved Azure image publisher, offer, SKU and version.' },
        [ordered]@{ field = 'network_and_security'; status = 'NEEDS_REVIEW'; action = 'Confirm network, NSG, IP, backup and security requirements before deployment.' }
    )

    $jsonObject = [ordered]@{
        schema_version = '1.0'
        source = [ordered]@{
            document_type = 'VM Build Sheet'
            excel_file_name = [System.IO.Path]::GetFileName($ExcelPath)
            service_request_no = $sheetRequestId
            build_requestor_name = $requestorName
            build_requestor_email = $requestorEmail
            application_server_owner_name = if ($ownerName) { $ownerName } else { $null }
            application_owner_email = if ($ownerEmail) { $ownerEmail } else { $null }
            application_team_dl = if ($teamDl) { $teamDl } else { $null }
        }
        approval = [ordered]@{
            status = 'NOT_VERIFIED'
            note = 'Approval has not been verified by this automation.'
        }
        shared_configuration = [ordered]@{
            subscription_name = $vms[0].subscription_name
            location_display_name = $vms[0].region_display_name
            resource_group = $vms[0].resource_group
            vnet_name = $vms[0].vnet_name
            subnet_name = $vms[0].subnet_name
            environment = $tags['Environment']
            tags = $tags
            access = $access
            domain = $domain
        }
        vms = @($vms)
        checklist = @($checklist)
        validation = [ordered]@{
            status = 'BASIC_VALIDATION_PASSED'
            errors = @()
            notes = @($validationNotes)
            deployment_ready = $false
        }
        generated_utc = [DateTime]::UtcNow.ToString('o')
    }

    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    $outputPath = Join-Path $OutputFolder "$RequestId.json"
    $json = $jsonObject | ConvertTo-Json -Depth 20
    # Parse once before writing to verify JSON syntax.
    $null = $json | ConvertFrom-Json -ErrorAction Stop
    Set-Content -LiteralPath $outputPath -Value $json -Encoding utf8

    Write-Host ''
    Write-Host '===================================='
    Write-Host 'JSON GENERATION COMPLETED'
    Write-Host '===================================='
    Write-Host "Request ID: $RequestId"
    Write-Host "VM count: $($vms.Count)"
    Write-Host "Checklist items: $($checklist.Count)"
    Write-Host 'Validation: BASIC_VALIDATION_PASSED'
    Write-Host 'Deployment ready: False (approval and Azure mappings are not verified)'
    Write-Host "Output file: $outputPath"
}
finally {
    if ($null -ne $package) {
        $package.Dispose()
    }
}
