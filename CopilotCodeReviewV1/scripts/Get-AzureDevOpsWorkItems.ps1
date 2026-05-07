<#
.SYNOPSIS
    Retrieves full details for work items linked to a pull request.

.DESCRIPTION
    This script uses the Azure DevOps REST API to fetch detailed information
    for one or more work items by their IDs. It extracts key fields such as
    title, description, acceptance criteria, and repro steps, and writes them
    to a structured text file for use as context in code reviews.

.PARAMETER Token
    Required. Authentication token for Azure DevOps. Can be a PAT or OAuth token.

.PARAMETER AuthType
    Optional. The type of authentication to use. Valid values: 'Basic' (for PAT) or 'Bearer' (for OAuth/System.AccessToken).
    Default is 'Basic'.

.PARAMETER CollectionUri
    Required. The Azure DevOps collection URI (e.g., 'https://dev.azure.com/myorg' or 'https://tfs.contoso.com/tfs/DefaultCollection').

.PARAMETER Project
    Required. The Azure DevOps project name.

.PARAMETER WorkItemIds
    Required. Comma-separated list of work item IDs to retrieve (e.g., '123,456,789').

.PARAMETER OutputFile
    Optional. Path to write the output to a file. If not specified, output is only written to the console.

.EXAMPLE
    .\Get-AzureDevOpsWorkItems.ps1 -Token "your-pat" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -WorkItemIds "123,456"
    Retrieves details for work items #123 and #456.

.EXAMPLE
    .\Get-AzureDevOpsWorkItems.ps1 -Token "oauth-token" -AuthType "Bearer" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -WorkItemIds "123" -OutputFile "C:\output\work-items.txt"
    Retrieves work item details using OAuth authentication and writes to a file.

.NOTES
    Author: Little Fort Software
    Date: March 2026
    Requires: PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Authentication token for Azure DevOps (PAT or OAuth token)")]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    [Parameter(Mandatory = $false, HelpMessage = "Authentication type: 'Basic' for PAT, 'Bearer' for OAuth")]
    [ValidateSet("Basic", "Bearer")]
    [string]$AuthType = "Basic",

    [Parameter(Mandatory = $true, HelpMessage = "Azure DevOps collection URI (e.g., https://dev.azure.com/myorg)")]
    [ValidateNotNullOrEmpty()]
    [string]$CollectionUri,

    [Parameter(Mandatory = $true, HelpMessage = "Azure DevOps project name")]
    [ValidateNotNullOrEmpty()]
    [string]$Project,

    [Parameter(Mandatory = $true, HelpMessage = "Comma-separated list of work item IDs (e.g., '123,456')")]
    [ValidateNotNullOrEmpty()]
    [string]$WorkItemIds,

    [Parameter(Mandatory = $false, HelpMessage = "Output file path to write results to")]
    [string]$OutputFile
)

#region Helper Functions

function Write-Output-Line {
    param(
        [string]$Message = "",
        [string]$ForegroundColor = "White",
        [switch]$NoNewline
    )

    if ($script:OutputToFile) {
        if ($NoNewline) {
            $script:OutputBuilder.Append($Message) | Out-Null
        }
        else {
            $script:OutputBuilder.AppendLine($Message) | Out-Null
        }
    }

    # Sanitize for Azure Pipelines: prevent ##vso[ and ##[ from being interpreted as logging commands
    $sanitized = $Message -replace '(?m)^##', ' ##'

    if ($NoNewline) {
        Write-Host $sanitized -ForegroundColor $ForegroundColor -NoNewline
    }
    else {
        Write-Host $sanitized -ForegroundColor $ForegroundColor
    }
}

function Get-AuthorizationHeader {
    param(
        [string]$Token,
        [string]$AuthType = "Basic"
    )

    if ($AuthType -eq "Bearer") {
        return @{
            Authorization  = "Bearer $Token"
            "Content-Type" = "application/json"
        }
    }
    else {
        $base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Token"))
        return @{
            Authorization  = "Basic $base64Auth"
            "Content-Type" = "application/json"
        }
    }
}

function Invoke-AzureDevOpsApi {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "Get"
    )

    try {
        $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ErrorAction Stop
        return $response
    }
    catch {
        $statusCode = $null
        $errorDetail = $null

        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorDetail = $_.ErrorDetails.Message
        }

        # Build a descriptive error message with all available context
        $baseMsg = "Azure DevOps API error"
        if ($statusCode) {
            $baseMsg += " (HTTP $statusCode)"
        }
        $baseMsg += " calling $Method $Uri"

        if ($statusCode -eq 401) {
            Write-Error "$baseMsg — Authentication failed. Please verify your token is valid and has work item read permissions. API response: $errorDetail"
        }
        elseif ($statusCode -eq 404) {
            Write-Error "$baseMsg — Work item(s) not found. Please verify the IDs are correct. API response: $errorDetail"
        }
        elseif ($statusCode) {
            Write-Error "$baseMsg — API response: $errorDetail"
        }
        else {
            Write-Error "$baseMsg — $($_.Exception.Message)"
        }
        return $null
    }
}

function ConvertFrom-Html {
    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    # Convert <br> tags to newlines
    $text = $Html -replace '<br\s*/?>', "`n"
    # Convert block-level tags to newlines
    $text = $text -replace '</(p|div|li|tr|h[1-6])>', "`n"
    $text = $text -replace '<(p|div|li|tr|h[1-6])[^>]*>', ""
    # Strip all remaining HTML tags
    $text = $text -replace '<[^>]+>', ''
    # Decode HTML entities
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    # Collapse multiple consecutive blank lines into one
    $text = $text -replace "(\r?\n\s*){3,}", "`n`n"
    # Trim leading/trailing whitespace
    $text = $text.Trim()

    return $text
}

#endregion

#region Main Logic

# Initialize output handling
$script:OutputToFile = -not [string]::IsNullOrEmpty($OutputFile)
$script:OutputBuilder = [System.Text.StringBuilder]::new()

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$apiVersion = "api-version=7.1"

# Validate and clean up work item IDs
$ids = ($WorkItemIds -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }

if ($ids.Count -eq 0) {
    Write-Warning "No valid work item IDs provided."
    exit 0
}

# Cap at 200 IDs (API batch limit)
if ($ids.Count -gt 200) {
    Write-Warning "More than 200 work items linked. Only the first 200 will be fetched."
    $ids = $ids | Select-Object -First 200
}

$idsParam = $ids -join ','

Write-Host "`nRetrieving details for work item(s): $idsParam..." -ForegroundColor Cyan

# Batch fetch work items
$workItemsUrl = "$CollectionUri/$Project/_apis/wit/workitems?ids=$idsParam&`$expand=all&$apiVersion"
$response = Invoke-AzureDevOpsApi -Uri $workItemsUrl -Headers $headers

if ($null -eq $response -or $null -eq $response.value -or $response.value.Count -eq 0) {
    Write-Warning "No work item details could be retrieved."
    exit 0
}

Write-Host "Retrieved $($response.value.Count) work item(s)." -ForegroundColor Green

# Display results
Write-Output-Line (("=" * 80)) -ForegroundColor DarkGray
Write-Output-Line "LINKED WORK ITEM DETAILS" -ForegroundColor Green
Write-Output-Line ("=" * 80) -ForegroundColor DarkGray

foreach ($wi in $response.value) {
    $fields = $wi.fields
    $wiType = $fields.'System.WorkItemType'
    $wiTitle = $fields.'System.Title'
    $wiState = $fields.'System.State'
    $wiDescription = ConvertFrom-Html $fields.'System.Description'
    $wiAcceptanceCriteria = ConvertFrom-Html $fields.'Microsoft.VSTS.Common.AcceptanceCriteria'
    $wiReproSteps = ConvertFrom-Html $fields.'Microsoft.VSTS.TCM.ReproSteps'

    Write-Output-Line "`n[Work Item #$($wi.id) - $wiType]" -ForegroundColor Yellow
    Write-Output-Line "  Title:           $wiTitle"
    Write-Output-Line "  State:           $wiState"

    if (-not [string]::IsNullOrWhiteSpace($wiDescription)) {
        Write-Output-Line "`n  Description:"
        # Indent each line of the description
        $wiDescription -split "`n" | ForEach-Object {
            Write-Output-Line "    $_"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($wiAcceptanceCriteria)) {
        Write-Output-Line "`n  Acceptance Criteria:"
        $wiAcceptanceCriteria -split "`n" | ForEach-Object {
            Write-Output-Line "    $_"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($wiReproSteps)) {
        Write-Output-Line "`n  Repro Steps:"
        $wiReproSteps -split "`n" | ForEach-Object {
            Write-Output-Line "    $_"
        }
    }
}

Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray

# Write to output file if specified
if ($script:OutputToFile) {
    try {
        $outputDir = Split-Path -Parent $OutputFile
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        $script:OutputBuilder.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "`nOutput written to: $OutputFile" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to write output file: $_"
    }
}

#endregion
