<#
.SYNOPSIS
    Updates the status of a comment thread in an Azure DevOps pull request.

.DESCRIPTION
    This script uses the Azure DevOps REST API to update the status of an existing
    comment thread in a pull request. This is useful for resolving or closing threads
    when issues have been addressed.

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
    Required. The pull request ID containing the comment thread.

.PARAMETER ThreadId
    Required. The ID of the comment thread to update.

.PARAMETER Status
    Required. The new status for the thread. Valid values: Active, Fixed, WontFix, Closed, Pending.

.EXAMPLE
    .\Update-AzureDevOpsPRCommentStatus.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -ThreadId 456 -Status "Fixed"
    Marks thread #456 as Fixed on pull request #123.

.EXAMPLE
    .\Update-AzureDevOpsPRCommentStatus.ps1 -Token $env:AZUREDEVOPS_TOKEN -AuthType "Bearer" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -ThreadId 456 -Status "Closed"
    Closes thread #456 using OAuth authentication.

.NOTES
    Author: Little Fort Software
    Date: January 2026
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

    [Parameter(Mandatory = $true, HelpMessage = "Thread ID to update")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$ThreadId,

    [Parameter(Mandatory = $true, HelpMessage = "New status for the thread")]
    [ValidateSet("Active", "Fixed", "WontFix", "Closed", "Pending")]
    [string]$Status
)

#region Helper Functions

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
        [string]$Method = "Get",
        [object]$Body = $null
    )
    
    try {
        $params = @{
            Uri         = $Uri
            Headers     = $Headers
            Method      = $Method
            ErrorAction = "Stop"
        }
        
        if ($null -ne $Body) {
            $params.Body = $Body | ConvertTo-Json -Depth 10
        }
        
        $response = Invoke-RestMethod @params
        return $response
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $errorMessage = $_.ErrorDetails.Message
        
        if ($statusCode -eq 401) {
            Write-Error "Authentication failed. Please verify your token is valid and has appropriate permissions."
        }
        elseif ($statusCode -eq 404) {
            Write-Error "Resource not found. Please verify the organization, project, repository, PR ID, and Thread ID."
        }
        elseif ($statusCode -eq 400) {
            Write-Error "Bad request: $errorMessage"
        }
        else {
            Write-Error "API request failed: $errorMessage (Status: $statusCode)"
        }
        return $null
    }
}

function Get-ThreadStatusValue {
    param([string]$StatusName)
    
    switch ($StatusName) {
        "Active"   { return 1 }
        "Fixed"    { return 2 }
        "WontFix"  { return 3 }
        "Closed"   { return 4 }
        "Pending"  { return 5 }
        default    { return 1 }
    }
}

function Format-DateForDisplay {
    param([string]$DateString)
    
    if ([string]::IsNullOrEmpty($DateString)) {
        return "N/A"
    }
    
    try {
        $date = [DateTime]::Parse($DateString)
        return $date.ToString("yyyy-MM-dd HH:mm")
    }
    catch {
        return $DateString
    }
}

#endregion

#region Main Logic

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$baseUrl = "https://dev.azure.com/$Organization/$Project/_apis/git/repositories/$Repository/pullrequests/$Id"
$apiVersion = "api-version=7.1"

# First, verify the thread exists
Write-Host "`nVerifying thread #$ThreadId exists on pull request #$Id..." -ForegroundColor Cyan
$threadUrl = "$baseUrl/threads/$ThreadId`?$apiVersion"
$existingThread = Invoke-AzureDevOpsApi -Uri $threadUrl -Headers $headers

if ($null -eq $existingThread) {
    Write-Error "Could not find thread #$ThreadId on pull request #$Id."
    exit 1
}

# Display current thread information
$currentStatus = switch ($existingThread.status) {
    "active"   { "Active" }
    "fixed"    { "Fixed" }
    "wontFix"  { "Won't Fix" }
    "closed"   { "Closed" }
    "pending"  { "Pending" }
    "byDesign" { "By Design" }
    default    { $existingThread.status }
}

Write-Host "Found thread #$ThreadId" -ForegroundColor Green
Write-Host "  Current Status: $currentStatus" -ForegroundColor $(if ($currentStatus -eq "Active") { "Yellow" } else { "Gray" })

if ($existingThread.comments -and $existingThread.comments.Count -gt 0) {
    $firstComment = $existingThread.comments[0]
    Write-Host "  Author: $($firstComment.author.displayName)" -ForegroundColor DarkGray
    Write-Host "  Posted: $(Format-DateForDisplay $firstComment.publishedDate)" -ForegroundColor DarkGray
    
    # Show file context if available
    if ($existingThread.threadContext -and $existingThread.threadContext.filePath) {
        $filePath = $existingThread.threadContext.filePath
        $lineInfo = ""
        if ($existingThread.threadContext.rightFileStart) {
            $lineInfo = " (Line $($existingThread.threadContext.rightFileStart.line))"
        }
        elseif ($existingThread.threadContext.leftFileStart) {
            $lineInfo = " (Line $($existingThread.threadContext.leftFileStart.line))"
        }
        Write-Host "  File: $filePath$lineInfo" -ForegroundColor DarkCyan
    }
}

# Check if status is already set to the requested value
if ($currentStatus -eq $Status) {
    Write-Host "`nThread #$ThreadId is already marked as '$Status'. No update needed." -ForegroundColor Yellow
    exit 0
}

# Update the thread status
Write-Host "`nUpdating thread status to '$Status'..." -ForegroundColor Cyan

$updateUrl = "$baseUrl/threads/$ThreadId`?$apiVersion"
$body = @{
    status = Get-ThreadStatusValue -StatusName $Status
}

$result = Invoke-AzureDevOpsApi -Uri $updateUrl -Headers $headers -Method "Patch" -Body $body

if ($null -ne $result) {
    Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
    Write-Host "THREAD STATUS UPDATED SUCCESSFULLY" -ForegroundColor Green
    Write-Host ("=" * 60) -ForegroundColor DarkGray
    Write-Host "`n  Thread ID:      #$ThreadId"
    Write-Host "  Pull Request:   #$Id"
    Write-Host "  Previous Status: $currentStatus" -ForegroundColor DarkGray
    Write-Host "  New Status:      $Status" -ForegroundColor Green
    Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
    
    # Provide link to the PR
    $webUrl = "https://dev.azure.com/$Organization/$Project/_git/$Repository/pullrequest/$Id"
    Write-Host "`nView PR: $webUrl" -ForegroundColor Cyan
}
else {
    Write-Error "Failed to update thread status."
    exit 1
}

#endregion
