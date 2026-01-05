<#
.SYNOPSIS
    Sets a status on a pull request in Azure DevOps.

.DESCRIPTION
    This script uses the Azure DevOps REST API to create or update a status check on a pull request.
    This allows the PR to be used with branch policies that require specific status checks to pass.

.PARAMETER Token
    Required. Authentication token for Azure DevOps. Can be a PAT or OAuth token.

.PARAMETER AuthType
    Optional. The type of authentication to use. Valid values: 'Basic' (for PAT) or 'Bearer' (for OAuth/System.AccessToken).
    Default is 'Basic'.

.PARAMETER Organization
    Required. The Azure DevOps organization name.

.PARAMETER Project
    Required. The Azure DevOps project name.

.PARAMETER Repository
    Required. The repository name where the pull request exists.

.PARAMETER Id
    Required. The pull request ID to set status on.

.PARAMETER State
    Required. The state of the status check. Valid values: succeeded, failed, error, pending, notSet, notApplicable.

.PARAMETER Description
    Optional. A description of the status check.

.PARAMETER TargetUrl
    Optional. A URL to provide more details about the status check.

.PARAMETER Genre
    Optional. The genre/category for this status check. Combined with Context as 'genre/context'.
    Default is 'copilot'.

.PARAMETER Context
    Optional. The name/identifier for this status check. Combined with Genre as 'genre/context'.
    Default is 'code review'.

.EXAMPLE
    .\Set-AzureDevOpsPRStatus.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -State "succeeded" -Description "Code review passed"
    Sets a succeeded status on pull request #123 with default context 'copilot/code review'.

.EXAMPLE
    .\Set-AzureDevOpsPRStatus.ps1 -Token "oauth-token" -AuthType "Bearer" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -State "failed" -Description "Issues found during review" -Genre "custom" -Context "security scan"
    Sets a failed status using OAuth authentication with context 'custom/security scan'.

.NOTES
    Author: Little Fort Software
    Date: January 2026
    Requires: PowerShell 5.1 or later
    
    API Documentation:
    https://learn.microsoft.com/en-us/rest/api/azure/devops/git/pull-request-statuses/create
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Authentication token for Azure DevOps (PAT or OAuth token)")]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    [Parameter(Mandatory = $false, HelpMessage = "Authentication type: 'Basic' for PAT, 'Bearer' for OAuth")]
    [ValidateSet("Basic", "Bearer")]
    [string]$AuthType = "Basic",

    [Parameter(Mandatory = $true, HelpMessage = "Azure DevOps organization name")]
    [ValidateNotNullOrEmpty()]
    [string]$Organization,

    [Parameter(Mandatory = $true, HelpMessage = "Azure DevOps project name")]
    [ValidateNotNullOrEmpty()]
    [string]$Project,

    [Parameter(Mandatory = $true, HelpMessage = "Repository name")]
    [ValidateNotNullOrEmpty()]
    [string]$Repository,

    [Parameter(Mandatory = $true, HelpMessage = "Pull request ID")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Id,

    [Parameter(Mandatory = $true, HelpMessage = "Status state")]
    [ValidateSet("succeeded", "failed", "error", "pending", "notSet", "notApplicable")]
    [string]$State,

    [Parameter(Mandatory = $false, HelpMessage = "Description of the status")]
    [string]$Description = "",

    [Parameter(Mandatory = $false, HelpMessage = "URL for more details")]
    [string]$TargetUrl = "",

    [Parameter(Mandatory = $false, HelpMessage = "Status genre/category")]
    [string]$Genre = "copilot",

    [Parameter(Mandatory = $false, HelpMessage = "Status context/name")]
    [string]$Context = "code review"
)

#region Helper Functions

function Get-AuthorizationHeader {
    param(
        [string]$Token,
        [string]$AuthType = "Basic"
    )
    
    if ($AuthType -eq "Bearer") {
        # OAuth/System.AccessToken - use Bearer authentication
        return @{
            Authorization = "Bearer $Token"
        }
    }
    else {
        # PAT - use Basic authentication
        $base64Token = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Token"))
        return @{
            Authorization = "Basic $base64Token"
        }
    }
}

function Invoke-AzureDevOpsApi {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers,
        [object]$Body
    )
    
    try {
        $params = @{
            Method      = $Method
            Uri         = $Uri
            Headers     = $Headers
            ContentType = "application/json"
        }
        
        if ($Body) {
            $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
        }
        
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        Write-Error "API request failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Error "Response: $responseBody"
        }
        throw
    }
}

#endregion

#region Main Script Logic

try {
    Write-Host "Setting PR status check..." -ForegroundColor Cyan
    Write-Host "  Organization: $Organization" -ForegroundColor Gray
    Write-Host "  Project: $Project" -ForegroundColor Gray
    Write-Host "  Repository: $Repository" -ForegroundColor Gray
    Write-Host "  PR ID: $Id" -ForegroundColor Gray
    Write-Host "  Context: $Genre/$Context" -ForegroundColor Gray
    Write-Host "  State: $State" -ForegroundColor Gray
    Write-Host "  Description: $Description" -ForegroundColor Gray
    
    # Build the API URL
    # https://learn.microsoft.com/en-us/rest/api/azure/devops/git/pull-request-statuses/create
    $apiUrl = "https://dev.azure.com/$Organization/$Project/_apis/git/repositories/$Repository/pullRequests/$Id/statuses?api-version=7.1"
    
    # Get authorization headers
    $headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
    
    # Build the status payload
    $statusPayload = @{
        state = $State
        description = $Description
        context = @{
            name = $Context
            genre = $Genre
        }
    }
    
    # Add target URL if provided
    if ($TargetUrl) {
        $statusPayload.targetUrl = $TargetUrl
    }
    
    # Create the status
    Write-Host "Posting status to Azure DevOps..." -ForegroundColor Cyan
    $result = Invoke-AzureDevOpsApi -Method "POST" -Uri $apiUrl -Headers $headers -Body $statusPayload
    
    Write-Host "✓ Status check set successfully!" -ForegroundColor Green
    Write-Host "  Status ID: $($result.id)" -ForegroundColor Gray
    
    exit 0
}
catch {
    Write-Error "Failed to set PR status: $($_.Exception.Message)"
    exit 1
}

#endregion
