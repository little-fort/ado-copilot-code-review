<#
.SYNOPSIS
    Records that a pull request iteration was reviewed, for incremental re-reviews.

.DESCRIPTION
    Writes one pull request property, <StateKeyPrefix>.Iteration.<Iteration>, marking that
    iteration as reviewed by the review configuration the prefix identifies (issue #67). Get-AzureDevOpsPRChanges.ps1
    reads these records on later runs to skip an iteration that was already reviewed, or to focus
    the review on what changed since.

    Every reviewed iteration gets its own property, so concurrent runs never overwrite each
    other's records and the write is idempotent (safe to retry). The task calls this script only
    after a review run that meets every recording condition (see shouldRecordReviewState in
    index.ts).

.PARAMETER Token
    Required. Authentication token for Azure DevOps. Can be a PAT or OAuth token.

.PARAMETER AuthType
    Optional. 'Basic' (for PAT) or 'Bearer' (for OAuth/System.AccessToken). Default is 'Basic'.

.PARAMETER CollectionUri
    Required. The Azure DevOps collection URI (e.g., 'https://dev.azure.com/myorg').

.PARAMETER Project
    Required. The Azure DevOps project name.

.PARAMETER Repository
    Required. The repository name where the pull request exists.

.PARAMETER Id
    Required. The pull request ID.

.PARAMETER StateKeyPrefix
    Required. Property-name prefix identifying the review configuration.

.PARAMETER Iteration
    Required. The pull request iteration that was reviewed.

.PARAMETER SourceCommit
    Required. The 40-character source commit SHA of that iteration.

.PARAMETER BuildId
    Optional. ID of the build that ran the review (Build.BuildId). Without it the record holds no
    build details (a run outside a build pipeline).

.PARAMETER BuildNumber
    Optional. Build number of that build (Build.BuildNumber); used only with -BuildId.

.PARAMETER BuildProject
    Optional. Project of that build (System.TeamProject), for the link to it; used only with -BuildId.

.EXAMPLE
    .\Set-AzureDevOpsPRReviewState.ps1 -Token "oauth-token" -AuthType "Bearer" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123 -StateKeyPrefix "CopilotCodeReview.v1.0123456789abcdef" -Iteration 4 -SourceCommit "<40-hex SHA>"
    Records iteration 4 of PR #123 as reviewed.

.NOTES
    Author: Little Fort Software
    Date: November 2026
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

    [Parameter(Mandatory = $true, HelpMessage = "Repository name")]
    [ValidateNotNullOrEmpty()]
    [string]$Repository,

    [Parameter(Mandatory = $true, HelpMessage = "Pull request ID")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Id,

    [Parameter(Mandatory = $true, HelpMessage = "Property-name prefix identifying the review configuration")]
    [ValidatePattern('^[A-Za-z0-9._-]{1,200}$')]
    [string]$StateKeyPrefix,

    [Parameter(Mandatory = $true, HelpMessage = "The reviewed pull request iteration")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Iteration,

    [Parameter(Mandatory = $true, HelpMessage = "Source commit SHA of the reviewed iteration")]
    [ValidatePattern('^[0-9a-fA-F]{40}$')]
    [string]$SourceCommit,

    [Parameter(Mandatory = $false, HelpMessage = "ID of the build that ran the review (Build.BuildId)")]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$BuildId = 0,

    [Parameter(Mandatory = $false, HelpMessage = "Build number of that build (Build.BuildNumber)")]
    [string]$BuildNumber = '',

    [Parameter(Mandatory = $false, HelpMessage = "Project of that build (System.TeamProject)")]
    [string]$BuildProject = ''
)

#region Helper Functions

function ConvertTo-SafeLogText {
    <#
        Neutralizes Azure Pipelines logging-command markers in text that came
        from outside this script (API responses, PR and work item content)
        before it is written to the pipeline log (issues #47, #69).
    #>
    param([string]$Text)

    return $Text -replace '##(?=(?:vso)?\[)', '# #'
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
    <#
        Calls the Azure DevOps REST API, retrying transient failures (throttling,
        server errors, network errors) with exponential backoff. The request this
        script sends replaces a single property with a fixed value, so replaying
        it after an ambiguous failure cannot cause harm. On final failure, prints
        ##[error] diagnostics (status and response body) and throws.
    #>
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "Get",
        [string]$Body = $null,
        [string]$ContentType = $null,
        [int]$MaxAttempts = 3
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Uri         = $Uri
                Headers     = $Headers
                Method      = $Method
                ErrorAction = "Stop"
            }
            if ($Body) {
                $params.Body = $Body
            }
            if ($ContentType) {
                $params.ContentType = $ContentType
            }
            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            $responseBody = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $responseBody = $_.ErrorDetails.Message
            }

            $isTransient = (-not $statusCode) -or $statusCode -eq 429 -or $statusCode -ge 500
            if ($isTransient -and $attempt -lt $MaxAttempts) {
                $delaySeconds = [math]::Pow(2, $attempt)
                $statusText = if ($statusCode) { "HTTP $statusCode" } else { ConvertTo-SafeLogText $_.Exception.Message }
                Write-Host "Transient Azure DevOps API failure ($statusText) on attempt $attempt of $MaxAttempts. Retrying in $delaySeconds seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            Write-Host "##[error]Azure DevOps API call failed: $Method $Uri"
            if ($statusCode) {
                Write-Host "##[error]HTTP status: $statusCode"
            }
            if ($responseBody) {
                Write-Host "##[error]API response: $(ConvertTo-SafeLogText $responseBody)"
            }
            else {
                Write-Host "##[error]$(ConvertTo-SafeLogText $_.Exception.Message)"
            }
            if ($statusCode -eq 401 -or $statusCode -eq 403) {
                Write-Host "##[error]Recording review state needs permission to update the pull request. If using the System Access Token, ensure the Build Service identity has 'Contribute to pull requests' on the repository."
            }
            throw
        }
    }
}

#endregion

#region Main Logic

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
# The properties endpoint takes a JSON Patch document, not plain JSON
$headers.Remove("Content-Type")

$propertiesUrl = "$CollectionUri/$Project/_apis/git/repositories/$Repository/pullRequests/$Id/properties?api-version=7.1"
$propertyName = "$($StateKeyPrefix).Iteration.$Iteration"

# The record (see .DESCRIPTION). Readers check every build detail's shape, so
# values that wouldn't pass (too long, line breaks) are left out here.
$record = [ordered]@{ sourceCommit = $SourceCommit.ToLowerInvariant() }
if ($BuildId -gt 0) {
    $record.buildId = $BuildId
    if ($BuildNumber -and $BuildNumber.Length -le 255 -and $BuildNumber -match '^[^\p{C}]+$') {
        $record.buildNumber = $BuildNumber
    }
    if ($BuildProject -and $BuildProject.Length -le 64 -and $BuildProject -match '^[^\p{C}]+$') {
        $record.buildProject = $BuildProject
    }
}
$record.recordedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
$recordValue = ConvertTo-Json -Compress -InputObject $record

# -InputObject keeps the one-operation patch a JSON array (piping a single
# object to ConvertTo-Json would produce a bare object, which the API rejects).
# 'replace' adds the property when it does not exist yet.
$patch = ConvertTo-Json -Compress -Depth 3 -InputObject @(
    [ordered]@{
        op    = 'replace'
        path  = "/$propertyName"
        value = $recordValue
    }
)

Write-Host "`nRecording iteration #$Iteration of pull request #$Id as reviewed..." -ForegroundColor Cyan
try {
    Invoke-AzureDevOpsApi -Uri $propertiesUrl -Headers $headers -Method "Patch" -Body $patch -ContentType "application/json-patch+json" | Out-Null
}
catch {
    Write-Host "##[error]Could not record iteration #$Iteration of pull request #$Id as reviewed."
    exit 1
}

Write-Host "Recorded iteration #$Iteration (source commit $($SourceCommit.Substring(0, 8))) as reviewed: $propertyName" -ForegroundColor Green
exit 0

#endregion
