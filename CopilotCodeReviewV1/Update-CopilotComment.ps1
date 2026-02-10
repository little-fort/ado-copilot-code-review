<#
.SYNOPSIS
    Updates an existing comment thread or comment on an Azure DevOps pull request.

.DESCRIPTION
    This script is used by GitHub Copilot to update previously-created comment threads
    or individual comments. It can update thread status (e.g., resolve a thread) and/or
    update the content of a specific comment within a thread. The script is designed to
    fail silently to avoid blocking the code review process if an update fails.

.PARAMETER ThreadId
    Required. The ID of the thread to update.

.PARAMETER Status
    Optional. The new status for the thread. Valid values: Active, Fixed, WontFix, Closed, Pending.
    If not specified and no other updates are provided, defaults to 'Fixed'.

.PARAMETER CommentId
    Optional. The ID of a specific comment within the thread to update.
    Required when updating comment content.

.PARAMETER Content
    Optional. New content for the comment. Requires CommentId to be specified.
    Supports markdown formatting.

.EXAMPLE
    .\Update-CopilotComment.ps1 -ThreadId 123 -Status Fixed
    Marks thread #123 as resolved/fixed.

.EXAMPLE
    .\Update-CopilotComment.ps1 -ThreadId 123 -CommentId 456 -Content "Updated feedback text"
    Updates the content of comment #456 in thread #123.

.EXAMPLE
    .\Update-CopilotComment.ps1 -ThreadId 123 -Status Fixed -CommentId 456 -Content "Issue resolved"
    Updates both the thread status and the comment content.

.NOTES
    Author: Little Fort Software
    Date: February 2026
    Requires: PowerShell 5.1 or later
    
    Environment Variables Used:
    - AZUREDEVOPS_TOKEN: Authentication token (PAT or OAuth)
    - AZUREDEVOPS_AUTH_TYPE: 'Basic' for PAT, 'Bearer' for OAuth
    - ORGANIZATION: Azure DevOps organization name
    - PROJECT: Azure DevOps project name
    - REPOSITORY: Repository name
    - PRID: Pull request ID
    
    This script fails silently by design. Errors are suppressed to avoid
    interrupting the Copilot review workflow.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Thread ID to update")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$ThreadId,

    [Parameter(Mandatory = $false, HelpMessage = "New status for the thread")]
    [ValidateSet("Active", "Fixed", "WontFix", "Closed", "Pending")]
    [string]$Status,

    [Parameter(Mandatory = $false, HelpMessage = "Comment ID to update (for content updates)")]
    [int]$CommentId,

    [Parameter(Mandatory = $false, HelpMessage = "New content for the comment")]
    [string]$Content
)

# Shared URL helpers
. "$PSScriptRoot\AzureDevOpsUrl.ps1"

# Wrap entire script in try/catch for silent failure
try {
    # Validate that at least one update is requested
    if ([string]::IsNullOrEmpty($Status) -and [string]::IsNullOrEmpty($Content)) {
        # Default to Fixed status if nothing specified
        $Status = "Fixed"
    }

    # Validate that Content requires CommentId
    if (-not [string]::IsNullOrEmpty($Content) -and $CommentId -le 0) {
        Write-Host "Note: Content update requires CommentId parameter" -ForegroundColor DarkGray
        exit 0
    }

    # Read credentials from environment variables
    $token = ${env:AZUREDEVOPS_TOKEN}
    $authType = ${env:AZUREDEVOPS_AUTH_TYPE}
    $organization = ${env:ORGANIZATION}
    $project = ${env:PROJECT}
    $repository = ${env:REPOSITORY}
    $prId = ${env:PRID}

    # Validate required environment variables
    if ([string]::IsNullOrEmpty($token) -or 
        [string]::IsNullOrEmpty($project) -or 
        [string]::IsNullOrEmpty($repository) -or 
        [string]::IsNullOrEmpty($prId)) {
        # Missing required env vars - exit silently
        exit 0
    }

    # Default auth type if not specified
    if ([string]::IsNullOrEmpty($authType)) {
        $authType = "Basic"
    }

    # Build authorization header
    if ($authType -eq "Bearer") {
        $headers = @{
            Authorization  = "Bearer $token"
            "Content-Type" = "application/json"
        }
    }
    else {
        $base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$token"))
        $headers = @{
            Authorization  = "Basic $base64Auth"
            "Content-Type" = "application/json"
        }
    }

    $baseUrls = Get-AzureDevOpsBaseUrls -Project $project -Organization $organization -Silent
    if ($null -eq $baseUrls) {
        exit 0
    }
    $baseUrl = $baseUrls.ApiBaseUrl

    # Update thread status if specified
    if (-not [string]::IsNullOrEmpty($Status)) {
        # Map status to API value (lowercase)
        $statusMap = @{
            "Active"  = "active"
            "Fixed"   = "fixed"
            "WontFix" = "wontFix"
            "Closed"  = "closed"
            "Pending" = "pending"
        }
        $apiStatus = $statusMap[$Status]

        $threadUri = "$baseUrl/git/repositories/$repository/pullrequests/$prId/threads/$ThreadId`?api-version=7.1"
        $threadBody = @{ status = $apiStatus } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri $threadUri -Headers $headers -Method Patch -Body $threadBody -ErrorAction Stop
        Write-Host "Thread #$ThreadId status updated to '$Status'" -ForegroundColor Green
    }

    # Update comment content if specified
    if (-not [string]::IsNullOrEmpty($Content) -and $CommentId -gt 0) {
        $commentUri = "$baseUrl/git/repositories/$repository/pullrequests/$prId/threads/$ThreadId/comments/$CommentId`?api-version=7.1"
        $commentBody = @{ content = $Content } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri $commentUri -Headers $headers -Method Patch -Body $commentBody -ErrorAction Stop
        Write-Host "Comment #$CommentId in thread #$ThreadId content updated" -ForegroundColor Green
    }
}
catch {
    # Silent failure - do not output error or set non-zero exit code
    # This ensures Copilot workflow continues even if update fails
    Write-Host "Note: Could not update thread #$ThreadId (this is not critical)" -ForegroundColor DarkGray
}

# Always exit with success
exit 0
