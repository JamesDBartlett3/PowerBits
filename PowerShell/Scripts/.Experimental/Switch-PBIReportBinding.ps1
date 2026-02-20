<#
.SYNOPSIS
    Rebinds a Power BI report to a different semantic model, supporting cross-workspace scenarios.

.DESCRIPTION
    Uses the Power BI REST API (Reports - Rebind Report In Group) to directly bind a report
    to a target semantic model. If the semantic model is in a different workspace, Power BI
    will automatically create a shared dataset reference in the report's workspace.

    Authentication follows the FabricPS-PBIP pattern using the Az.Accounts module.
    Supports interactive login, service principal, and existing Az session reuse.

    This performs a DIRECT binding — it does NOT create a composite model.

.PARAMETER ReportWorkspaceName
    Display name of the workspace containing the report.

.PARAMETER ReportWorkspaceId
    GUID of the workspace containing the report. Alternative to ReportWorkspaceName.

.PARAMETER ReportName
    Display name of the report to rebind.

.PARAMETER ReportId
    GUID of the report to rebind. Alternative to ReportName.

.PARAMETER DatasetWorkspaceName
    Display name of the workspace containing the target semantic model.
    If omitted, assumes same workspace as the report.

.PARAMETER DatasetWorkspaceId
    GUID of the workspace containing the target semantic model. Alternative to DatasetWorkspaceName.

.PARAMETER DatasetName
    Display name of the target semantic model.

.PARAMETER DatasetId
    GUID of the target semantic model. Alternative to DatasetName.

.PARAMETER ServicePrincipalId
    App (client) ID for service principal authentication.

.PARAMETER ServicePrincipalSecret
    Client secret for service principal authentication.

.PARAMETER TenantId
    Azure AD tenant ID (required for service principal auth).

.PARAMETER WhatIf
    Shows what would happen without actually performing the rebind.

.EXAMPLE
    # Rebind by names (same workspace)
    .\Switch-PBIReportBinding.ps1 -ReportWorkspaceName "Sales" -ReportName "Monthly Report" -DatasetName "Sales Model v2"

.EXAMPLE
    # Rebind cross-workspace by names
    .\Switch-PBIReportBinding.ps1 -ReportWorkspaceName "Reports WS" -ReportName "Executive Dashboard" `
        -DatasetWorkspaceName "Models WS" -DatasetName "Enterprise Model"

.EXAMPLE
    # Rebind by GUIDs with service principal
    .\Switch-PBIReportBinding.ps1 -ReportWorkspaceId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -ReportId "11111111-2222-3333-4444-555555555555" `
        -DatasetId "66666666-7777-8888-9999-000000000000" `
        -ServicePrincipalId "app-id" -ServicePrincipalSecret "secret" -TenantId "tenant-id"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # --- Report location ---
    [Parameter(ParameterSetName = "ByName")]
    [Parameter(ParameterSetName = "ByNameSP")]
    [string]$ReportWorkspaceName,

    [Parameter(ParameterSetName = "ById")]
    [Parameter(ParameterSetName = "ByIdSP")]
    [guid]$ReportWorkspaceId,

    [Parameter(ParameterSetName = "ByName")]
    [Parameter(ParameterSetName = "ByNameSP")]
    [string]$ReportName,

    [Parameter(ParameterSetName = "ById")]
    [Parameter(ParameterSetName = "ByIdSP")]
    [guid]$ReportId,

    # --- Target semantic model location ---
    [Parameter(ParameterSetName = "ByName")]
    [Parameter(ParameterSetName = "ByNameSP")]
    [string]$DatasetWorkspaceName,

    [Parameter(ParameterSetName = "ById")]
    [Parameter(ParameterSetName = "ByIdSP")]
    [guid]$DatasetWorkspaceId,

    [Parameter(ParameterSetName = "ByName")]
    [Parameter(ParameterSetName = "ByNameSP")]
    [string]$DatasetName,

    [Parameter(ParameterSetName = "ById")]
    [Parameter(ParameterSetName = "ByIdSP")]
    [guid]$DatasetId,

    # --- Service Principal auth ---
    [Parameter(ParameterSetName = "ByNameSP", Mandatory)]
    [Parameter(ParameterSetName = "ByIdSP", Mandatory)]
    [string]$ServicePrincipalId,

    [Parameter(ParameterSetName = "ByNameSP", Mandatory)]
    [Parameter(ParameterSetName = "ByIdSP", Mandatory)]
    [string]$ServicePrincipalSecret,

    [Parameter(ParameterSetName = "ByNameSP", Mandatory)]
    [Parameter(ParameterSetName = "ByIdSP", Mandatory)]
    [string]$TenantId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Authentication ──────────────────────────────────────────────────────

function Get-PwrBtsAccessToken {
    <#
    .SYNOPSIS
        Acquires a bearer token for the Power BI REST API.
        Pattern borrowed from FabricPS-PBIP (Set-FabricAuthToken).
    #>
    param(
        [string]$ServicePrincipalId,
        [string]$ServicePrincipalSecret,
        [string]$TenantId
    )

    $resourceUrl = "https://analysis.windows.net/powerbi/api"

    # Ensure Az.Accounts is available
    if (-not (Get-Module -Name "Az.Accounts" -ListAvailable)) {
        throw "Az.Accounts module is required. Install it with: Install-Module Az.Accounts -Scope CurrentUser"
    }
    Import-Module Az.Accounts -ErrorAction Stop

    if ($ServicePrincipalId) {
        Write-Host "Authenticating with service principal..." -ForegroundColor Cyan
        $secureSecret = ConvertTo-SecureString $ServicePrincipalSecret -AsPlainText -Force
        $credential = [PSCredential]::new($ServicePrincipalId, $secureSecret)
        Connect-AzAccount -ServicePrincipal -Credential $credential -TenantId $TenantId -ErrorAction Stop | Out-Null
    }
    else {
        # Reuse existing session or prompt interactive login
        $context = Get-AzContext
        if (-not $context) {
            Write-Host "No active Az session found. Launching interactive login..." -ForegroundColor Cyan
            Connect-AzAccount -ErrorAction Stop | Out-Null
        }
        else {
            Write-Host "Reusing existing Az session for '$($context.Account.Id)'." -ForegroundColor Cyan
        }
    }

    $tokenResult = Get-AzAccessToken -ResourceUrl $resourceUrl -ErrorAction Stop
    # Az.Accounts >= 5.x returns a SecureString; older versions return plain text
    if ($tokenResult.Token -is [System.Security.SecureString]) {
        $token = $tokenResult.Token | ConvertFrom-SecureString -AsPlainText
    }
    else {
        $token = $tokenResult.Token
    }

    if ([string]::IsNullOrWhiteSpace($token)) {
        throw "Failed to acquire an access token."
    }

    return $token
}

#endregion

#region REST API Helper ─────────────────────────────────────────────────────

function Invoke-PwrBtsApiRequest {
    <#
    .SYNOPSIS
        Thin wrapper around Invoke-RestMethod for Power BI REST API calls.
        Mirrors the Invoke-FabricAPIRequest pattern from FabricPS-PBIP.
    #>
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet("Get", "Post", "Patch", "Put", "Delete")]
        [string]$Method = "Get",
        [object]$Body,
        [int]$MaxRetries = 3
    )

    $baseUrl = "https://api.powerbi.com/v1.0/myorg"
    $fullUrl = "$baseUrl/$($Uri.TrimStart('/'))"

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    $splat = @{
        Uri         = $fullUrl
        Method      = $Method
        Headers     = $headers
        ErrorAction = "Stop"
    }

    if ($Body) {
        $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
        $splat["Body"] = $jsonBody
    }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-RestMethod @splat
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # Retry on 429 (throttle) or transient 5xx
            if ($attempt -le $MaxRetries -and ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600))) {
                $retryAfter = 5
                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                $wait = [Math]::Max($retryAfter, [Math]::Pow(2, $attempt))
                Write-Warning "HTTP $statusCode on attempt $attempt/$MaxRetries. Retrying in ${wait}s..."
                Start-Sleep -Seconds $wait
                continue
            }

            # Try to extract a meaningful error body
            $errorDetail = $_.ErrorDetails.Message
            if ($errorDetail) {
                try {
                    $parsed = $errorDetail | ConvertFrom-Json
                    $errorDetail = $parsed.error.message ?? $parsed.error.code ?? $errorDetail
                }
                catch { }
            }

            throw "Power BI API error (HTTP $statusCode) on ${Method} ${fullUrl}: $errorDetail`n$($_.Exception.Message)"
        }
    }
}

#endregion

#region Lookup Helpers ──────────────────────────────────────────────────────

function Resolve-PwrBtsWorkspaceId {
    param(
        [Parameter(Mandatory)][string]$Token,
        [string]$WorkspaceName,
        [guid]$WorkspaceId
    )

    if ($WorkspaceId -and $WorkspaceId -ne [guid]::Empty) {
        return $WorkspaceId.ToString()
    }

    if ([string]::IsNullOrWhiteSpace($WorkspaceName)) {
        throw "Either a workspace name or workspace ID must be provided."
    }

    Write-Host "  Looking up workspace '$WorkspaceName'..." -ForegroundColor DarkGray
    # Use filter to find by exact name
    $encodedName = [Uri]::EscapeDataString($WorkspaceName)
    $result = Invoke-PwrBtsApiRequest -Token $Token -Uri "groups?`$filter=name eq '$encodedName'"

    $foundItems = @($result.value | Where-Object { $_.name -eq $WorkspaceName })

    if ($foundItems.Count -eq 0) {
        throw "Workspace '$WorkspaceName' not found. Check the name and your access permissions."
    }
    if ($foundItems.Count -gt 1) {
        throw "Multiple workspaces found with name '$WorkspaceName'. Use -ReportWorkspaceId / -DatasetWorkspaceId instead."
    }

    Write-Host "  -> Workspace ID: $($foundItems[0].id)" -ForegroundColor DarkGray
    return $foundItems[0].id
}


function Resolve-PwrBtsReportId {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$WorkspaceId,
        [string]$ReportName,
        [guid]$ReportId
    )

    if ($ReportId -and $ReportId -ne [guid]::Empty) {
        return $ReportId.ToString()
    }

    if ([string]::IsNullOrWhiteSpace($ReportName)) {
        throw "Either a report name or report ID must be provided."
    }

    Write-Host "  Looking up report '$ReportName' in workspace $WorkspaceId..." -ForegroundColor DarkGray
    $result = Invoke-PwrBtsApiRequest -Token $Token -Uri "groups/$WorkspaceId/reports"

    $matchedReports = @($result.value | Where-Object { $_.name -eq $ReportName })

    if ($matchedReports.Count -eq 0) {
        throw "Report '$ReportName' not found in workspace $WorkspaceId."
    }
    if ($matchedReports.Count -gt 1) {
        Write-Warning "Multiple reports named '$ReportName' found. Using the first match."
    }

    $report = $matchedReports[0]
    Write-Host "  -> Report ID: $($report.id)  (currently bound to dataset: $($report.datasetId))" -ForegroundColor DarkGray
    return $report.id
}


function Resolve-PwrBtsDatasetId {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$WorkspaceId,
        [string]$DatasetName,
        [guid]$DatasetId
    )

    if ($DatasetId -and $DatasetId -ne [guid]::Empty) {
        return $DatasetId.ToString()
    }

    if ([string]::IsNullOrWhiteSpace($DatasetName)) {
        throw "Either a dataset/semantic model name or ID must be provided."
    }

    Write-Host "  Looking up semantic model '$DatasetName' in workspace $WorkspaceId..." -ForegroundColor DarkGray
    $result = Invoke-PwrBtsApiRequest -Token $Token -Uri "groups/$WorkspaceId/datasets"

    $foundDatasets = @($result.value | Where-Object { $_.name -eq $DatasetName })

    if ($foundDatasets.Count -eq 0) {
        throw "Semantic model '$DatasetName' not found in workspace $WorkspaceId."
    }
    if ($foundDatasets.Count -gt 1) {
        Write-Warning "Multiple semantic models named '$DatasetName' found. Using the first match."
    }

    Write-Host "  -> Dataset ID: $($foundDatasets[0].id)" -ForegroundColor DarkGray
    return $foundDatasets[0].id
}

#endregion

#region Main ────────────────────────────────────────────────────────────────

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  Switch-PBIReportBinding" -ForegroundColor Cy
Write-Host "================================================================`n" -ForegroundColor Cyan

# 1. Authenticate
$token = Get-PwrBtsAccessToken `
    -ServicePrincipalId $ServicePrincipalId `
    -ServicePrincipalSecret $ServicePrincipalSecret `
    -TenantId $TenantId

# 2. Resolve report workspace
Write-Host "[1/4] Resolving report workspace..." -ForegroundColor Yellow
$resolvedReportWsId = Resolve-PwrBtsWorkspaceId -Token $token -WorkspaceName $ReportWorkspaceName -WorkspaceId $ReportWorkspaceId

# 3. Resolve report
Write-Host "[2/4] Resolving report..." -ForegroundColor Yellow
$resolvedReportId = Resolve-PwrBtsReportId -Token $token -WorkspaceId $resolvedReportWsId -ReportName $ReportName -ReportId $ReportId

# 4. Resolve dataset workspace (defaults to report workspace if not specified)
Write-Host "[3/4] Resolving target semantic model workspace..." -ForegroundColor Yellow
$targetDatasetWsId = if ($DatasetWorkspaceName -or ($DatasetWorkspaceId -and $DatasetWorkspaceId -ne [guid]::Empty)) {
    Resolve-PwrBtsWorkspaceId -Token $token -WorkspaceName $DatasetWorkspaceName -WorkspaceId $DatasetWorkspaceId
}
else {
    Write-Host "  -> Using report workspace (no separate dataset workspace specified)." -ForegroundColor DarkGray
    $resolvedReportWsId
}

# 5. Resolve dataset
Write-Host "[4/4] Resolving target semantic model..." -ForegroundColor Yellow
$resolvedDatasetId = Resolve-PwrBtsDatasetId -Token $token -WorkspaceId $targetDatasetWsId -DatasetName $DatasetName -DatasetId $DatasetId

# 6. Get current report details for confirmation
$reportDetails = Invoke-PwrBtsApiRequest -Token $token -Uri "groups/$resolvedReportWsId/reports/$resolvedReportId"

$isCrossWorkspace = $resolvedReportWsId -ne $targetDatasetWsId

Write-Host "`n----------------------------------------------------------------" -ForegroundColor White
Write-Host "  Rebind Summary" -ForegroundColor White
Write-Host "----------------------------------------------------------------" -ForegroundColor White
Write-Host "  Report:           $($reportDetails.name)" -ForegroundColor White
Write-Host "  Report ID:        $resolvedReportId" -ForegroundColor White
Write-Host "  Report Workspace: $resolvedReportWsId" -ForegroundColor White
Write-Host "  Current Dataset:  $($reportDetails.datasetId)" -ForegroundColor White
Write-Host "  Target Dataset:   $resolvedDatasetId" -ForegroundColor Green
if ($isCrossWorkspace) {
    Write-Host "  Cross-Workspace:  YES (shared dataset ref will be created)" -ForegroundColor Magenta
}
Write-Host "----------------------------------------------------------------`n" -ForegroundColor White

if ($reportDetails.datasetId -eq $resolvedDatasetId) {
    Write-Host "Report is already bound to the target dataset. Nothing to do." -ForegroundColor Yellow
    return
}

# 7. Perform the rebind
$rebindBody = @{ datasetId = $resolvedDatasetId }

if ($PSCmdlet.ShouldProcess("Report '$($reportDetails.name)' [$resolvedReportId]", "Rebind to dataset $resolvedDatasetId")) {
    Write-Host "Executing rebind..." -ForegroundColor Cyan
    Invoke-PwrBtsApiRequest -Token $token `
        -Uri "groups/$resolvedReportWsId/reports/$resolvedReportId/Rebind" `
        -Method Post `
        -Body $rebindBody

    Write-Host "[OK] Rebind successful!" -ForegroundColor Green

    # Verify
    Write-Host "Verifying..." -ForegroundColor Cyan
    $updatedReport = Invoke-PwrBtsApiRequest -Token $token -Uri "groups/$resolvedReportWsId/reports/$resolvedReportId"

    if ($updatedReport.datasetId -eq $resolvedDatasetId) {
        Write-Host "[OK] Verified: Report is now bound to dataset $($updatedReport.datasetId)" -ForegroundColor Green
    }
    else {
        Write-Warning "Verification mismatch: report shows datasetId '$($updatedReport.datasetId)' — expected '$resolvedDatasetId'. Check the Power BI Service lineage view."
    }
}

#endregion
