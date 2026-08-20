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

.PARAMETER CollectionUri
    Required. The Azure DevOps collection URI (e.g., 'https://dev.azure.com/myorg' or 'https://tfs.contoso.com/tfs/DefaultCollection').

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
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "This looks good!"
    Creates a new comment thread on pull request #123 using PAT authentication.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "oauth-token" -AuthType "Bearer" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "This looks good!"
    Creates a new comment thread using OAuth/System.AccessToken authentication.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "I agree" -ThreadId 456
    Replies to an existing thread #456 on pull request #123.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "Consider async" -FilePath "/src/Program.cs" -StartLine 42
    Creates an inline comment on line 42 of Program.cs.

.EXAMPLE
    .\Add-AzureDevOpsPRComment.ps1 -Token "your-pat" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123 -Comment "Refactor this" -FilePath "/src/Program.cs" -StartLine 42 -EndLine 50 -IterationId 3
    Creates an inline comment spanning lines 42-50, anchored to iteration 3 of the PR.

.NOTES
    Author: Little Fort Software
    Date: December 2025
    Requires: PowerShell 5.1 or later
    
    If an inline comment fails (e.g., line no longer exists in the diff), the script will
    automatically fall back to posting a generic PR comment with the file path and line
    information appended to the comment text.

    Exit codes: 0 on success, 1 on any failure. API failures are reported on stdout
    with ##[error] markers (including the HTTP status and response body) so they are
    visible in pipeline logs and to the calling agent. Transient failures (HTTP 429,
    5xx, network errors) are retried with exponential backoff before failing.
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
        [object]$Body = $null,
        [int]$MaxAttempts = 3,
        # For non-idempotent requests (POST): a probe that checks whether the
        # failed request was actually committed server-side. Must return the
        # created resource (used as the response) or $null if nothing landed.
        [scriptblock]$VerifyCompleted = $null
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

            if ($null -ne $Body) {
                $params.Body = $Body | ConvertTo-Json -Depth 10
            }

            $response = Invoke-RestMethod @params
            return $response
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

            # Throttling, server errors, and network failures are worth retrying;
            # other 4xx responses are deterministic and fail immediately.
            $isThrottled = $statusCode -eq 429
            $isTransient = (-not $statusCode) -or $isThrottled -or $statusCode -ge 500

            # POST is not idempotent: a 5xx or network error can occur AFTER the
            # server committed the write, so blindly replaying the request could
            # create a duplicate comment. Before retrying (and before giving up
            # on the final attempt), probe whether the write actually landed and
            # reuse it if so. 429 means the request was rejected unprocessed, so
            # no probe is needed there.
            if ($isTransient -and -not $isThrottled -and $Method -eq 'Post' -and $VerifyCompleted) {
                $committed = $null
                try {
                    $committed = & $VerifyCompleted
                }
                catch {
                    # The probe itself failed (likely the same outage). Fall
                    # through to the normal retry path: a rare duplicate comment
                    # is preferable to silently dropping review feedback.
                    Write-Host "Could not verify whether the failed $Method request was committed; proceeding with retry." -ForegroundColor Yellow
                }
                if ($null -ne $committed) {
                    Write-Host "The failed $Method request was committed by the server despite the error; using the existing resource instead of retrying." -ForegroundColor Yellow
                    return $committed
                }
            }

            if ($isTransient -and $attempt -lt $MaxAttempts) {
                $delaySeconds = [math]::Pow(2, $attempt)
                $statusText = if ($statusCode) { "HTTP $statusCode" } else { $_.Exception.Message }
                Write-Host "Transient Azure DevOps API failure ($statusText) on attempt $attempt of $MaxAttempts. Retrying in $delaySeconds seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $delaySeconds
                continue
            }

            # Emit full diagnostics on stdout: stderr (Write-Error) is dropped in
            # several invocation paths, which previously made these failures
            # impossible to diagnose from the pipeline log (issue #56). The
            # ##[error] prefix renders red in Azure Pipelines logs.
            Write-Host "##[error]Azure DevOps API call failed: $Method $Uri"
            if ($statusCode) {
                Write-Host "##[error]HTTP status: $statusCode"
            }
            if ($responseBody) {
                Write-Host "##[error]API response: $responseBody"
            }
            else {
                Write-Host "##[error]$($_.Exception.Message)"
            }

            if ($statusCode -eq 401) {
                Write-Host "##[error]Authentication failed. Please verify your token is valid and has appropriate permissions."
            }
            elseif ($statusCode -eq 403) {
                Write-Host "##[error]Permission denied. If using the System Access Token, ensure the Build Service identity has 'Contribute to pull requests' on the repository."
            }
            elseif ($statusCode -eq 404) {
                Write-Host "##[error]Resource not found. Please verify the organization, project, repository, and PR ID."
            }

            # Let the caller decide whether a failure is fatal (e.g. inline
            # comments fall back to a general comment before giving up)
            throw
        }
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
$baseUrl = "$CollectionUri/$Project/_apis/git/repositories/$Repository/pullrequests/$Id"
$apiVersion = "api-version=7.1"

# First, verify the PR exists
Write-Host "`nVerifying pull request #$Id exists..." -ForegroundColor Cyan
$prUrl = "$baseUrl`?$apiVersion"
try {
    $pr = Invoke-AzureDevOpsApi -Uri $prUrl -Headers $headers
}
catch {
    Write-Host "##[error]Could not retrieve pull request #$Id in repository '$Repository'."
    exit 1
}

Write-Host "Found PR: $($pr.title)" -ForegroundColor Green

if ($ThreadId -gt 0) {
    # Reply to existing thread
    Write-Host "`nReplying to thread #$ThreadId..." -ForegroundColor Cyan
    
    # Verify the thread exists
    $threadUrl = "$baseUrl/threads/$ThreadId`?$apiVersion"
    try {
        $existingThread = Invoke-AzureDevOpsApi -Uri $threadUrl -Headers $headers
    }
    catch {
        Write-Host "##[error]Could not retrieve thread #$ThreadId on pull request #$Id."
        exit 1
    }

    # Post reply to the thread
    $commentsUrl = "$baseUrl/threads/$ThreadId/comments?$apiVersion"
    $body = @{
        content       = $Comment
        parentCommentId = 0
        commentType   = 1  # Text comment
    }

    # Duplicate guard for ambiguous POST failures: re-read the thread and treat
    # a reply with the exact same text as the committed result
    $verifyReplyCommitted = {
        $thread = Invoke-AzureDevOpsApi -Uri $threadUrl -Headers $headers
        @($thread.comments) | Select-Object -Skip 1 | Where-Object { $_.content -eq $Comment } | Select-Object -First 1
    }

    try {
        $result = Invoke-AzureDevOpsApi -Uri $commentsUrl -Headers $headers -Method "Post" -Body $body -VerifyCompleted $verifyReplyCommitted
    }
    catch {
        Write-Host "##[error]Failed to post reply to thread #$ThreadId on pull request #$Id."
        exit 1
    }

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

    # Duplicate guard for ambiguous POST failures: list the PR's threads and
    # treat one whose first comment has the exact same text as the committed
    # result. Content equality is safe here because the probe only runs seconds
    # after our own POST failed ambiguously.
    $verifyThreadCommitted = {
        $threads = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers
        @($threads.value) | Where-Object {
            $_.comments -and @($_.comments)[0].content -eq $Comment
        } | Select-Object -First 1
    }

    # Attempt to post the comment
    try {
        $result = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers -Method "Post" -Body $body -VerifyCompleted $verifyThreadCommitted
    }
    catch {
        if ($isInlineComment) {
            $inlineCommentFailed = $true
            Write-Warning "Failed to post inline comment. Falling back to generic PR comment with file/line information appended."
        }
        else {
            Write-Host "##[error]Failed to create comment thread on pull request #$Id."
            exit 1
        }
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
        
        # Same duplicate guard as above, matching the fallback comment text
        $verifyFallbackCommitted = {
            $threads = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers
            @($threads.value) | Where-Object {
                $_.comments -and @($_.comments)[0].content -eq $fallbackComment
            } | Select-Object -First 1
        }

        Write-Host "Posting generic comment with file/line information..." -ForegroundColor Yellow
        try {
            $result = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers -Method "Post" -Body $fallbackBody -VerifyCompleted $verifyFallbackCommitted
        }
        catch {
            Write-Host "##[error]Fallback generic comment also failed for pull request #$Id."
            exit 1
        }
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

# Defensive: every failure path above should already have exited, but never
# report success without a confirmed API result (issue #56)
if ($null -eq $result) {
    Write-Host "##[error]The Azure DevOps API returned no result for the comment operation on pull request #$Id."
    exit 1
}

# Provide link to the PR
$webUrl = "$CollectionUri/$Project/_git/$Repository/pullrequest/$Id"
Write-Host "`nView PR: $webUrl" -ForegroundColor Cyan

# Explicit success exit so callers reading $LASTEXITCODE never see a stale value
exit 0

#endregion
