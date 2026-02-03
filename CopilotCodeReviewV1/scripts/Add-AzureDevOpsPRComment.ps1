<#
.SYNOPSIS
    Posts a comment to a pull request in Azure DevOps.

.DESCRIPTION
    This script uses the Azure DevOps REST API to add a comment to a pull request.
    It can either create a new comment thread or reply to an existing thread.
    Supports both general PR-level comments and file-specific inline comments.

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
    Required. The pull request ID to comment on.

.PARAMETER Comment
    Required. The comment text to post. Supports markdown formatting.

.PARAMETER ThreadId
    Optional. The ID of an existing thread to reply to. If not specified, a new thread is created.

.PARAMETER Status
    Optional. The status for a new thread. Valid values: Active, Fixed, WontFix, Closed, Pending.
    Default is 'Active'. Only applies when creating a new thread (not replying).

.PARAMETER FilePath
    Optional. File path for inline comment (e.g., '/src/MyProject/Program.cs').
    When provided with StartLine, creates an inline comment on the specified file.
    Path will be normalized to use forward slashes with a leading slash.

.PARAMETER StartLine
    Optional. Starting line number for inline comment (1-based, references the right/changed side of the diff).
    Required when FilePath is provided for inline comments.

.PARAMETER EndLine
    Optional. Ending line number for inline comment. Defaults to StartLine if not provided.

.PARAMETER IterationId
    Optional. Pull request iteration ID for inline comments. Helps anchor the comment to the correct diff version.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "This looks good!"
    Creates a new comment thread on pull request #123 using PAT authentication.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "oauth-token" -AuthType "Bearer" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "This looks good!"
    Creates a new comment thread using OAuth/System.AccessToken authentication.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "I agree" -ThreadId 456
    Replies to an existing thread #456 on pull request #123.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "Consider async" -FilePath "/src/Program.cs" -StartLine 42
    Creates an inline comment on line 42 of Program.cs.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "Refactor this" -FilePath "/src/Program.cs" -StartLine 42 -EndLine 50 -IterationId 3
    Creates an inline comment spanning lines 42-50, anchored to iteration 3 of the PR.

.NOTES
    Author: Little Fort Software
    Date: December 2025
    Requires: PowerShell 5.1 or later
    
    If an inline comment fails (e.g., line no longer exists in the diff), the script will
    automatically fall back to posting a generic PR comment with the file path and line
    information appended to the comment text.
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

    [Parameter(Mandatory = $true, HelpMessage = "Comment text to post")]
    [ValidateNotNullOrEmpty()]
    [string]$Comment,

    [Parameter(Mandatory = $false, HelpMessage = "Existing thread ID to reply to")]
    [int]$ThreadId,

    [Parameter(Mandatory = $false, HelpMessage = "Status for new thread")]
    [ValidateSet("Active", "Fixed", "WontFix", "Closed", "Pending")]
    [string]$Status = "Active",

    [Parameter(Mandatory = $false, HelpMessage = "File path for inline comment (e.g., '/src/MyProject/Program.cs')")]
    [string]$FilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Starting line number for inline comment")]
    [int]$StartLine,

    [Parameter(Mandatory = $false, HelpMessage = "Ending line number for inline comment")]
    [int]$EndLine,

    [Parameter(Mandatory = $false, HelpMessage = "Pull request iteration ID for inline comments")]
    [int]$IterationId
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
            Write-Error "Authentication failed. Please verify your PAT is valid and has appropriate permissions."
        }
        elseif ($statusCode -eq 404) {
            Write-Error "Resource not found. Please verify the organization, project, repository, and PR ID."
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

function Format-AzureDevOpsFilePath {
    param([string]$Path)
    
    # Normalize path separators to forward slashes
    $normalized = $Path -replace '\\', '/'
    
    # Ensure path starts with a forward slash
    if (-not $normalized.StartsWith('/')) {
        $normalized = '/' + $normalized
    }
    
    return $normalized
}

#endregion

#region Main Logic

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$baseUrl = "https://dev.azure.com/$Organization/$Project/_apis/git/repositories/$Repository/pullrequests/$Id"
$apiVersion = "api-version=7.1"

# First, verify the PR exists
Write-Host "`nVerifying pull request #$Id exists..." -ForegroundColor Cyan
$prUrl = "$baseUrl`?$apiVersion"
$pr = Invoke-AzureDevOpsApi -Uri $prUrl -Headers $headers

if ($null -eq $pr) {
    Write-Error "Could not find pull request #$Id in repository '$Repository'."
    exit 1
}

Write-Host "Found PR: $($pr.title)" -ForegroundColor Green

if ($ThreadId -gt 0) {
    # Reply to existing thread
    Write-Host "`nReplying to thread #$ThreadId..." -ForegroundColor Cyan
    
    # Verify the thread exists
    $threadUrl = "$baseUrl/threads/$ThreadId`?$apiVersion"
    $existingThread = Invoke-AzureDevOpsApi -Uri $threadUrl -Headers $headers
    
    if ($null -eq $existingThread) {
        Write-Error "Could not find thread #$ThreadId on pull request #$Id."
        exit 1
    }
    
    # Post reply to the thread
    $commentsUrl = "$baseUrl/threads/$ThreadId/comments?$apiVersion"
    $body = @{
        content       = $Comment
        parentCommentId = 0
        commentType   = 1  # Text comment
    }
    
    $result = Invoke-AzureDevOpsApi -Uri $commentsUrl -Headers $headers -Method "Post" -Body $body
    
    if ($null -ne $result) {
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
        Write-Host "COMMENT POSTED SUCCESSFULLY" -ForegroundColor Green
        Write-Host ("=" * 60) -ForegroundColor DarkGray
        Write-Host "`n  Thread ID:    #$ThreadId"
        Write-Host "  Comment ID:   #$($result.id)"
        Write-Host "  Author:       $($result.author.displayName)"
        Write-Host "  Posted:       $($result.publishedDate)"
        Write-Host "`n  Content:"
        Write-Host "  $Comment" -ForegroundColor White
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
    }
}
else {
    # Create new thread
    $threadsUrl = "$baseUrl/threads?$apiVersion"
    $isInlineComment = -not [string]::IsNullOrEmpty($FilePath) -and $StartLine -gt 0
    
    # Build the base body
    $body = @{
        comments = @(
            @{
                content     = $Comment
                commentType = 1  # Text comment
            }
        )
        status   = Get-ThreadStatusValue -StatusName $Status
    }
    
    # Add threadContext for inline comments
    if ($isInlineComment) {
        $normalizedPath = Format-AzureDevOpsFilePath -Path $FilePath
        $effectiveEndLine = if ($EndLine -gt 0) { $EndLine } else { $StartLine }
        
        Write-Host "`nCreating inline comment thread on $normalizedPath (Lines $StartLine-$effectiveEndLine)..." -ForegroundColor Cyan
        
        $body.threadContext = @{
            filePath       = $normalizedPath
            rightFileStart = @{
                line   = $StartLine
                offset = 1
            }
            rightFileEnd   = @{
                line   = $effectiveEndLine
                offset = 1
            }
        }
        
        # Add iteration context if available
        if ($IterationId -gt 0) {
            $body.pullRequestThreadContext = @{
                iterationContext = @{
                    firstComparingIteration = $IterationId
                    secondComparingIteration = $IterationId
                }
            }
        }
    } else {
        Write-Host "`nCreating new comment thread..." -ForegroundColor Cyan
    }
    
    $result = $null
    $inlineCommentFailed = $false
    
    # Attempt to post the comment
    try {
        $result = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers -Method "Post" -Body $body
    }
    catch {
        if ($isInlineComment) {
            $inlineCommentFailed = $true
            Write-Warning "Failed to post inline comment: $($_.Exception.Message)"
            Write-Warning "Falling back to generic PR comment with file/line information appended."
        }
        else {
            throw
        }
    }
    
    # Check if inline comment failed (result is null but was inline)
    if ($null -eq $result -and $isInlineComment -and -not $inlineCommentFailed) {
        $inlineCommentFailed = $true
        Write-Warning "Inline comment API returned no result. Falling back to generic PR comment."
    }
    
    # Fallback to generic comment if inline failed
    if ($inlineCommentFailed) {
        $normalizedPath = Format-AzureDevOpsFilePath -Path $FilePath
        $effectiveEndLine = if ($EndLine -gt 0) { $EndLine } else { $StartLine }
        
        # Append file/line info to the comment
        $lineInfo = if ($StartLine -eq $effectiveEndLine) { "Line $StartLine" } else { "Lines $StartLine-$effectiveEndLine" }
        $fallbackComment = $Comment + "`n`n**File:** ``$normalizedPath```n**$lineInfo**"
        
        $fallbackBody = @{
            comments = @(
                @{
                    content     = $fallbackComment
                    commentType = 1
                }
            )
            status   = Get-ThreadStatusValue -StatusName $Status
        }
        
        Write-Host "Posting generic comment with file/line information..." -ForegroundColor Yellow
        $result = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers -Method "Post" -Body $fallbackBody
    }
    
    if ($null -ne $result) {
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
        Write-Host "COMMENT THREAD CREATED SUCCESSFULLY" -ForegroundColor Green
        Write-Host ("=" * 60) -ForegroundColor DarkGray
        Write-Host "`n  Thread ID:    #$($result.id)"
        Write-Host "  Status:       $Status"
        if ($isInlineComment -and -not $inlineCommentFailed) {
            Write-Host "  Type:         Inline comment"
            Write-Host "  File:         $(Format-AzureDevOpsFilePath -Path $FilePath)"
            $effectiveEndLine = if ($EndLine -gt 0) { $EndLine } else { $StartLine }
            Write-Host "  Lines:        $StartLine-$effectiveEndLine"
        } else {
            Write-Host "  Type:         General comment"
        }
        Write-Host "  Comment ID:   #$($result.comments[0].id)"
        Write-Host "  Author:       $($result.comments[0].author.displayName)"
        Write-Host "  Posted:       $($result.comments[0].publishedDate)"
        Write-Host "`n  Content:"
        Write-Host "  $Comment" -ForegroundColor White
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
        
        Write-Host "`nTip: Use -ThreadId $($result.id) to reply to this thread." -ForegroundColor DarkGray
    }
}

# Provide link to the PR
$webUrl = "https://dev.azure.com/$Organization/$Project/_git/$Repository/pullrequest/$Id"
Write-Host "`nView PR: $webUrl" -ForegroundColor Cyan

#endregion
