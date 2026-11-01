<#
.SYNOPSIS
    Retrieves open pull requests from Azure DevOps using the REST API.

.DESCRIPTION
    This script queries the Azure DevOps REST API to retrieve pull request information.
    It can either list all open PRs matching specified criteria in a tabular format,
    or display detailed information for a specific PR when an ID is provided.

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
    Optional. Filter PRs by repository name. If not specified, PRs from all repositories are returned.

.PARAMETER Creator
    Optional. Filter PRs by creator's display name or email (partial match supported).

.PARAMETER Id
    Optional. Specific pull request ID to retrieve detailed information for.

.PARAMETER IncludeHumanThreads
    Optional. Also emit a HUMAN REVIEWER THREADS (JSON) section listing active threads opened by human
    reviewers, with their full text, author, and replies. The task passes this only when the review agent
    is allowed to resolve human threads (resolveThreads = all).

.EXAMPLE
    .\Get-AzureDevOpsPR.ps1 -Token "your-pat-token" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject"
    Lists all open pull requests in the project using PAT authentication.

.EXAMPLE
    .\Get-AzureDevOpsPR.ps1 -Token "oauth-token" -AuthType "Bearer" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject"
    Lists all open pull requests using OAuth/System.AccessToken authentication.

.EXAMPLE
    .\Get-AzureDevOpsPR.ps1 -Token "your-pat-token" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo"
    Lists all open pull requests in a specific repository.

.EXAMPLE
    .\Get-AzureDevOpsPR.ps1 -Token "your-pat-token" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Id 123
    Displays detailed information for pull request #123.

.EXAMPLE
    .\Get-AzureDevOpsPR.ps1 -Token "your-pat-token" -CollectionUri "https://tfs.contoso.com/tfs/DefaultCollection" -Project "myproject" -Id 123 -OutputFile "C:\output\pr-details.txt"
    Writes the pull request details to the specified file (on-prem example).

.NOTES
    Author: Little Fort Software
    Date: December 2025
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

    [Parameter(Mandatory = $false, HelpMessage = "Repository name (optional)")]
    [string]$Repository,

    [Parameter(Mandatory = $false, HelpMessage = "Filter by creator display name or email")]
    [string]$Creator,

    [Parameter(Mandatory = $false, HelpMessage = "Specific pull request ID")]
    [int]$Id,

    [Parameter(Mandatory = $false, HelpMessage = "Output file path to write results to")]
    [string]$OutputFile,

    [Parameter(Mandatory = $false, HelpMessage = "Include active human reviewer threads as JSON")]
    [switch]$IncludeHumanThreads
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

    # Sanitize the console copy only; the output file keeps the original text
    # for the review agent
    $sanitized = ConvertTo-SafeLogText $Message

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
            Authorization = "Bearer $Token"
            "Content-Type" = "application/json"
        }
    }
    else {
        $base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Token"))
        return @{
            Authorization = "Basic $base64Auth"
            "Content-Type" = "application/json"
        }
    }
}

function Invoke-AzureDevOpsApi {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )
    
    try {
        $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -ErrorAction Stop
        return $response
    }
    catch {
        $statusCode = $null
        $errorDetail = $null

        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errorDetail = ConvertTo-SafeLogText $_.ErrorDetails.Message
        }

        # Build a descriptive error message with all available context
        $baseMsg = "Azure DevOps API error"
        if ($statusCode) {
            $baseMsg += " (HTTP $statusCode)"
        }
        $baseMsg += " calling GET $Uri"

        if ($statusCode -eq 401) {
            Write-Error "$baseMsg — Authentication failed. Please verify your token is valid and has appropriate permissions. API response: $errorDetail"
        }
        elseif ($statusCode -eq 404) {
            Write-Error "$baseMsg — Resource not found. Please verify the organization, project, and repository names. API response: $errorDetail"
        }
        elseif ($statusCode) {
            Write-Error "$baseMsg — API response: $errorDetail"
        }
        else {
            Write-Error "$baseMsg — $(ConvertTo-SafeLogText $_.Exception.Message)"
        }
        return $null
    }
}

function Invoke-AzureDevOpsApiPaginated {
    param(
        [string]$BaseUri,
        [hashtable]$Headers,
        [int]$PageSize = 250,
        [int]$MaxResults = 0,
        [scriptblock]$StopCondition = $null
    )
    
    $allResults = @()
    $skip = 0
    $page = 1
    
    # Determine the separator for query parameters
    $separator = if ($BaseUri -match '\?') { '&' } else { '?' }
    
    do {
        if ($page -eq 1) {
            Write-Host "Fetching pull requests (page $page)..." -ForegroundColor DarkGray
        }
        else {
            Write-Host "Fetching pull requests (page $page, $($allResults.Count) retrieved so far)..." -ForegroundColor DarkGray
        }
        
        $paginatedUri = "$BaseUri$separator`$top=$PageSize&`$skip=$skip"
        $response = Invoke-AzureDevOpsApi -Uri $paginatedUri -Headers $Headers
        
        if ($null -eq $response -or $null -eq $response.value) {
            break
        }
        
        $returnedCount = $response.value.Count
        
        # Check stop condition after each page (for early termination)
        if ($null -ne $StopCondition) {
            $match = & $StopCondition $response.value
            if ($null -ne $match) {
                Write-Host "Found target on page $page." -ForegroundColor DarkGray
                return @{ value = @($match); count = 1; earlyTermination = $true }
            }
        }
        
        $allResults += $response.value
        $skip += $PageSize
        $page++
        
        # Check if we should stop due to MaxResults
        if ($MaxResults -gt 0 -and $allResults.Count -ge $MaxResults) {
            $allResults = $allResults | Select-Object -First $MaxResults
            break
        }
        
    } while ($returnedCount -eq $PageSize)  # Continue if we got a full page
    
    Write-Host "Retrieved $($allResults.Count) pull requests total." -ForegroundColor DarkGray
    return @{ value = $allResults; count = $allResults.Count }
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

function Get-BranchShortName {
    param([string]$RefName)
    
    if ([string]::IsNullOrEmpty($RefName)) {
        return "N/A"
    }
    
    return $RefName -replace "^refs/heads/", ""
}

function Get-ReviewersSummary {
    param($Reviewers)
    
    if ($null -eq $Reviewers -or $Reviewers.Count -eq 0) {
        return "None"
    }
    
    $reviewerInfo = $Reviewers | ForEach-Object {
        $vote = switch ($_.vote) {
            10 { "[Approved]" }
            5  { "[Approved with suggestions]" }
            0  { "[No vote]" }
            -5 { "[Waiting]" }
            -10 { "[Rejected]" }
            default { "[Unknown]" }
        }
        "$($_.displayName) $vote"
    }
    
    return $reviewerInfo -join "; "
}

function Test-AgentAuthoredThread {
    <#
        Returns $true when a comment thread was created by a previous automated
        review. Match by: Build Service author OR the agent attribution tag in
        the first comment. The tag check matters when comments are posted with
        a PAT — the author is then a human account, so only the tag identifies
        the thread as agent-created.

        The tag must be matched with -match and escaped brackets: in a -like
        pattern, [brackets] form a wildcard character class, so a pattern like
        '*[Generated by ...]*' would match nearly every human comment and leak
        human threads into the agent-thread JSON (found by PR #61 review).
    #>
    param($Thread)

    $firstComment = $Thread.comments[0]

    if (($firstComment.author.displayName -like "*Build Service*") -or
        ($firstComment.author.uniqueName -like "*Build Service*")) {
        return $true
    }

    return [bool]($firstComment.content -match '\[Generated by (GitHub Copilot|Claude Code)\]')
}

function Get-CommentPreviewLines {
    <#
        Returns the first $MaxLines non-blank, trimmed lines of a comment and
        how many non-blank lines were left out, for the prose thread summary.
        Only non-blank lines are displayed, so only those are counted — the
        "(N more lines)" note then matches what the reader actually missed.
    #>
    param(
        [string]$Content,
        [int]$MaxLines
    )

    $lines = @(($Content -split "`r?`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $shown = @($lines | Select-Object -First $MaxLines)
    return [pscustomobject]@{
        Lines   = $shown
        Omitted = $lines.Count - $shown.Count
    }
}

function ConvertTo-ThreadJsonObject {
    <#
        Builds the JSON shape of a thread section the review agent acts on
        (COPILOT COMMENT THREADS / HUMAN REVIEWER THREADS): the ids the comment
        scripts need, status, file position, the full first comment, and every
        reply. System messages and soft-deleted replies (which have no content)
        are left out. -IncludeAuthor adds the first comment's author, which
        matters for human threads; automated-review threads are all ours.
    #>
    param(
        $Thread,
        [switch]$IncludeAuthor
    )

    $firstComment = $Thread.comments | Select-Object -First 1

    $filePath = $null
    $startLine = $null
    if ($Thread.threadContext -and $Thread.threadContext.filePath) {
        $filePath = $Thread.threadContext.filePath
        if ($Thread.threadContext.rightFileStart) {
            $startLine = $Thread.threadContext.rightFileStart.line
        }
        elseif ($Thread.threadContext.leftFileStart) {
            $startLine = $Thread.threadContext.leftFileStart.line
        }
    }

    $replies = @(@($Thread.comments) | Select-Object -Skip 1 | Where-Object {
        $_.commentType -ne 'system' -and -not $_.isDeleted
    } | ForEach-Object {
        [ordered]@{
            author        = $_.author.displayName
            content       = $_.content
            publishedDate = $_.publishedDate
        }
    })

    # Ordered so the JSON reads the same way on every run
    $result = [ordered]@{
        threadId  = $Thread.id
        commentId = $firstComment.id
        status    = $Thread.status
        filePath  = $filePath
        startLine = $startLine
    }
    if ($IncludeAuthor) {
        $result.author = $firstComment.author.displayName
    }
    $result.content = $firstComment.content
    $result.replies = $replies
    return $result
}

#endregion

#region Main Logic

# Initialize output handling
$script:OutputToFile = -not [string]::IsNullOrEmpty($OutputFile)
$script:OutputBuilder = [System.Text.StringBuilder]::new()

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$baseUrl = "$CollectionUri/$Project/_apis"
$apiVersion = "api-version=7.1"

# If a specific PR ID is provided, get detailed information
if ($Id -gt 0) {
    Write-Host "`nRetrieving details for Pull Request #$Id..." -ForegroundColor Cyan
    
    # First, we need to find the PR across repositories if repository is not specified
    if ([string]::IsNullOrEmpty($Repository)) {
        # Search for the PR across all repositories in the project (with pagination)
        $searchUrl = "$baseUrl/git/pullrequests?searchCriteria.status=all&$apiVersion"
        
        # Use early termination to stop as soon as we find the target PR
        $targetPrId = $Id
        $stopCondition = {
            param($results)
            $results | Where-Object { $_.pullRequestId -eq $targetPrId } | Select-Object -First 1
        }
        
        $allPRs = Invoke-AzureDevOpsApiPaginated -BaseUri $searchUrl -Headers $headers -StopCondition $stopCondition
        
        if ($null -eq $allPRs -or $null -eq $allPRs.value -or $allPRs.value.Count -eq 0) {
            Write-Warning "Pull Request #$Id not found in project '$Project'."
            exit 0
        }
        
        $targetPR = $allPRs.value | Where-Object { $_.pullRequestId -eq $Id } | Select-Object -First 1
        
        if ($null -eq $targetPR) {
            Write-Warning "Pull Request #$Id not found in project '$Project'."
            exit 0
        }
        
        $Repository = $targetPR.repository.name
    }
    
    # Get detailed PR information
    $prUrl = "$baseUrl/git/repositories/$Repository/pullrequests/$Id`?$apiVersion"
    $pr = Invoke-AzureDevOpsApi -Uri $prUrl -Headers $headers
    
    if ($null -eq $pr) {
        Write-Warning "Pull Request #$Id not found in repository '$Repository'."
        exit 0
    }
    
    # Get work items linked to the PR
    $workItemsUrl = "$baseUrl/git/repositories/$Repository/pullrequests/$Id/workitems?$apiVersion"
    $workItems = Invoke-AzureDevOpsApi -Uri $workItemsUrl -Headers $headers
    
    # Get PR iterations (commits info)
    $iterationsUrl = "$baseUrl/git/repositories/$Repository/pullrequests/$Id/iterations?$apiVersion"
    $iterations = Invoke-AzureDevOpsApi -Uri $iterationsUrl -Headers $headers
    
    # Get PR threads (comments)
    $threadsUrl = "$baseUrl/git/repositories/$Repository/pullrequests/$Id/threads?$apiVersion"
    $threads = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers
    
    # Display detailed information
    Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray
    Write-Output-Line "PULL REQUEST DETAILS" -ForegroundColor Green
    Write-Output-Line ("=" * 80) -ForegroundColor DarkGray
    
    Write-Output-Line "`n[Basic Information]" -ForegroundColor Yellow
    Write-Output-Line "  ID:              #$($pr.pullRequestId)"
    Write-Output-Line "  Title:           $($pr.title)"
    $statusColor = switch ($pr.status) {
        "active" { "Green" }
        "completed" { "Blue" }
        "abandoned" { "Red" }
        default { "White" }
    }
    Write-Output-Line "  Status:          $($pr.status.ToUpper())" -ForegroundColor $statusColor
    Write-Output-Line "  Repository:      $($pr.repository.name)"
    Write-Output-Line "  Source Branch:   $(Get-BranchShortName $pr.sourceRefName)"
    Write-Output-Line "  Target Branch:   $(Get-BranchShortName $pr.targetRefName)"
    Write-Output-Line "  Is Draft:        $($pr.isDraft)"
    Write-Output-Line "  Merge Status:    $($pr.mergeStatus)"
    
    Write-Output-Line "`n[People]" -ForegroundColor Yellow
    Write-Output-Line "  Created By:      $($pr.createdBy.displayName) <$($pr.createdBy.uniqueName)>"
    Write-Output-Line "  Created Date:    $(Format-DateForDisplay $pr.creationDate)"
    
    if ($pr.closedBy) {
        Write-Output-Line "  Closed By:       $($pr.closedBy.displayName)"
        Write-Output-Line "  Closed Date:     $(Format-DateForDisplay $pr.closedDate)"
    }
    
    Write-Output-Line "`n[Reviewers]" -ForegroundColor Yellow
    if ($pr.reviewers -and $pr.reviewers.Count -gt 0) {
        foreach ($reviewer in $pr.reviewers) {
            $voteDisplay = switch ($reviewer.vote) {
                10 { "Approved"; "Green" }
                5  { "Approved with suggestions"; "Yellow" }
                0  { "No vote"; "Gray" }
                -5 { "Waiting for author"; "Yellow" }
                -10 { "Rejected"; "Red" }
                default { "Unknown"; "White" }
            }
            $required = if ($reviewer.isRequired) { " (Required)" } else { "" }
            Write-Output-Line "  - $($reviewer.displayName)$required : $($voteDisplay[0])" -ForegroundColor $voteDisplay[1]
        }
    }
    else {
        Write-Output-Line "  No reviewers assigned"
    }
    
    Write-Output-Line "`n[Description]" -ForegroundColor Yellow
    if ([string]::IsNullOrEmpty($pr.description)) {
        Write-Output-Line "  (No description provided)"
    }
    else {
        $description = $pr.description -replace "`r`n", "`n" -replace "`n", "`n  "
        Write-Output-Line "  $description"
    }
    
    Write-Output-Line "`n[Iterations/Updates]" -ForegroundColor Yellow
    if ($iterations -and $iterations.value) {
        Write-Output-Line "  Total iterations: $($iterations.value.Count)"
        $lastIteration = $iterations.value | Select-Object -Last 1
        if ($lastIteration) {
            Write-Output-Line "  Last updated:     $(Format-DateForDisplay $lastIteration.updatedDate)"
        }
    }
    
    # Classify comment threads once (system-generated threads excluded); the
    # prose summary and the JSON sections below share these sets
    $commentThreads = @()
    $agentThreads = @()
    $humanActiveThreads = @()
    if ($threads -and $threads.value) {
        $commentThreads = @($threads.value | Where-Object {
            $_.comments -and
            $_.comments.Count -gt 0 -and
            $_.comments[0].commentType -ne "system"
        })
        # Threads created by a previous automated review (see Test-AgentAuthoredThread)
        $agentThreads = @($commentThreads | Where-Object { Test-AgentAuthoredThread -Thread $_ })
        if ($IncludeHumanThreads) {
            $humanActiveThreads = @($commentThreads | Where-Object {
                -not (Test-AgentAuthoredThread -Thread $_) -and $_.status -in @('active', 'pending')
            })
        }
    }

    # Threads printed in full in a JSON section get a one-line pointer in the
    # prose summary instead of a second copy of their text
    $agentThreadIds = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($t in $agentThreads) { [void]$agentThreadIds.Add([int]$t.id) }
    $humanThreadIds = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($t in $humanActiveThreads) { [void]$humanThreadIds.Add([int]$t.id) }

    Write-Output-Line "`n[Comments/Threads]" -ForegroundColor Yellow
    if ($threads -and $threads.value) {
        $activeThreads = $commentThreads | Where-Object { $_.status -eq "active" }
        $resolvedThreads = $commentThreads | Where-Object { $_.status -eq "fixed" -or $_.status -eq "closed" }
        
        Write-Output-Line "  Active threads:   $($activeThreads.Count)"
        Write-Output-Line "  Resolved threads: $($resolvedThreads.Count)"
        
        if ($commentThreads.Count -gt 0) {
            Write-Output-Line "`n  --- Top-Level Comments ---" -ForegroundColor DarkGray
            
            foreach ($thread in $commentThreads) {
                $firstComment = $thread.comments | Select-Object -First 1
                $threadStatus = switch ($thread.status) {
                    "active"   { @{ Text = "Active"; Color = "Yellow" } }
                    "fixed"    { @{ Text = "Resolved"; Color = "Green" } }
                    "closed"   { @{ Text = "Closed"; Color = "Green" } }
                    "wontFix"  { @{ Text = "Won't Fix"; Color = "DarkGray" } }
                    "pending"  { @{ Text = "Pending"; Color = "Cyan" } }
                    "byDesign" { @{ Text = "By Design"; Color = "DarkGray" } }
                    default    { @{ Text = $thread.status; Color = "White" } }
                }
                
                Write-Output-Line ""
                Write-Output-Line "  Thread #$($thread.id) [$($threadStatus.Text)]" -ForegroundColor $threadStatus.Color
                
                # Show file context if this is a file-level comment
                if ($thread.threadContext -and $thread.threadContext.filePath) {
                    $filePath = $thread.threadContext.filePath
                    $lineInfo = ""
                    if ($thread.threadContext.rightFileStart) {
                        $lineInfo = " (Line $($thread.threadContext.rightFileStart.line))"
                    }
                    elseif ($thread.threadContext.leftFileStart) {
                        $lineInfo = " (Line $($thread.threadContext.leftFileStart.line))"
                    }
                    Write-Output-Line "  File: $filePath$lineInfo" -ForegroundColor DarkCyan
                }
                
                Write-Output-Line "  Author: $($firstComment.author.displayName) | $(Format-DateForDisplay $firstComment.publishedDate)" -ForegroundColor DarkGray
                
                if ($agentThreadIds.Contains([int]$thread.id)) {
                    Write-Output-Line "    (automated review thread; full text in the COPILOT COMMENT THREADS (JSON) section below)" -ForegroundColor DarkGray
                }
                elseif ($humanThreadIds.Contains([int]$thread.id)) {
                    Write-Output-Line "    (full text and replies in the HUMAN REVIEWER THREADS (JSON) section below)" -ForegroundColor DarkGray
                }
                else {
                    # Resolved threads stay listed so the agent does not repeat
                    # existing feedback, but a short preview is enough for that
                    $maxLines = if ($thread.status -in @('fixed', 'closed', 'wontFix', 'byDesign')) { 5 } else { 30 }
                    $preview = Get-CommentPreviewLines -Content $firstComment.content -MaxLines $maxLines
                    foreach ($line in $preview.Lines) {
                        Write-Output-Line "    $line"
                    }
                    if ($preview.Omitted -gt 0) {
                        Write-Output-Line "    ... ($($preview.Omitted) more lines)" -ForegroundColor DarkGray
                    }
                }
                
                # Show reply count
                $replyCount = $thread.comments.Count - 1
                if ($replyCount -gt 0) {
                    Write-Output-Line "    [$replyCount $(if ($replyCount -eq 1) { 'reply' } else { 'replies' })]" -ForegroundColor DarkGray
                }
            }
        }
    }
    else {
        Write-Output-Line "  No comments"
    }
    
    Write-Output-Line "`n[Linked Work Items]" -ForegroundColor Yellow
    if ($workItems -and $workItems.value -and $workItems.value.Count -gt 0) {
        foreach ($wi in $workItems.value) {
            Write-Output-Line "  - #$($wi.id): $($wi.url)"
        }
    }
    else {
        Write-Output-Line "  No linked work items"
    }
    
    Write-Output-Line "`n[Links]" -ForegroundColor Yellow
    $webUrl = "$CollectionUri/$Project/_git/$($pr.repository.name)/pullrequest/$($pr.pullRequestId)"
    Write-Output-Line "  Web URL: $webUrl"
    
    Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray

    # Structured JSON for threads from previous automated reviews, which the
    # agent replies to and resolves. -InputObject keeps a single thread a
    # one-element array (piping one object to ConvertTo-Json yields a bare object).
    if ($agentThreads.Count -gt 0) {
        Write-Output-Line "`n=== COPILOT COMMENT THREADS (JSON) ===" -ForegroundColor Cyan
        $jsonThreads = @($agentThreads | ForEach-Object { ConvertTo-ThreadJsonObject -Thread $_ })
        Write-Output-Line (ConvertTo-Json -InputObject $jsonThreads -Compress -Depth 6)
        Write-Output-Line "=== END COPILOT COMMENT THREADS ===" -ForegroundColor Cyan
    }

    # Active threads from human reviewers, only when the agent may resolve them
    # (resolveThreads = all): it needs the full text and replies to judge whether
    # a concern was addressed or disputed, and the ids to reply and resolve
    if ($humanActiveThreads.Count -gt 0) {
        Write-Output-Line "`n=== HUMAN REVIEWER THREADS (JSON) ===" -ForegroundColor Cyan
        $jsonThreads = @($humanActiveThreads | ForEach-Object { ConvertTo-ThreadJsonObject -Thread $_ -IncludeAuthor })
        Write-Output-Line (ConvertTo-Json -InputObject $jsonThreads -Compress -Depth 6)
        Write-Output-Line "=== END HUMAN REVIEWER THREADS ===" -ForegroundColor Cyan
    }
}
else {
    # List all open PRs
    Write-Host "`nRetrieving open pull requests..." -ForegroundColor Cyan
    
    $pullRequests = @()
    
    if ([string]::IsNullOrEmpty($Repository)) {
        # Get PRs from all repositories in the project (with pagination)
        $prsUrl = "$baseUrl/git/pullrequests?searchCriteria.status=active&$apiVersion"
        $response = Invoke-AzureDevOpsApiPaginated -BaseUri $prsUrl -Headers $headers
        
        if ($null -eq $response) {
            exit 1
        }
        
        $pullRequests = $response.value
    }
    else {
        # Get PRs from specific repository (with pagination)
        $prsUrl = "$baseUrl/git/repositories/$Repository/pullrequests?searchCriteria.status=active&$apiVersion"
        $response = Invoke-AzureDevOpsApiPaginated -BaseUri $prsUrl -Headers $headers
        
        if ($null -eq $response) {
            exit 1
        }
        
        $pullRequests = $response.value
    }
    
    # Filter by creator if specified
    if (-not [string]::IsNullOrEmpty($Creator)) {
        $pullRequests = $pullRequests | Where-Object {
            $_.createdBy.displayName -like "*$Creator*" -or
            $_.createdBy.uniqueName -like "*$Creator*"
        }
    }
    
    if ($pullRequests.Count -eq 0) {
        Write-Output-Line "`nNo open pull requests found matching the specified criteria." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Output-Line "`nFound $($pullRequests.Count) open pull request(s):`n" -ForegroundColor Green
    
    # Create table output
    $tableData = $pullRequests | ForEach-Object {
        [PSCustomObject]@{
            "ID"         = $_.pullRequestId
            "Title"      = if ($_.title.Length -gt 50) { $_.title.Substring(0, 47) + "..." } else { $_.title }
            "Repository" = $_.repository.name
            "Source"     = Get-BranchShortName $_.sourceRefName
            "Target"     = Get-BranchShortName $_.targetRefName
            "Created By" = $_.createdBy.displayName
            "Created"    = Format-DateForDisplay $_.creationDate
            "Draft"      = if ($_.isDraft) { "Yes" } else { "No" }
        }
    }
    
    $tableOutput = $tableData | Format-Table -AutoSize -Wrap | Out-String
    Write-Host (ConvertTo-SafeLogText $tableOutput)
    if ($script:OutputToFile) {
        $script:OutputBuilder.AppendLine($tableOutput) | Out-Null
    }
    
    Write-Output-Line "`nTip: Use -Id <number> parameter to view detailed information for a specific PR." -ForegroundColor DarkGray
}

# Write to output file if specified
if ($script:OutputToFile) {
    try {
        $outputDir = Split-Path -Parent $OutputFile
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        $script:OutputBuilder.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "`nOutput written to: $OutputFile" -ForegroundColor Green

        # Write linked work item IDs to a separate file for downstream consumption
        if ($workItems -and $workItems.value -and $workItems.value.Count -gt 0) {
            $workItemIds = ($workItems.value | ForEach-Object { $_.id }) -join ","
            $workItemIdsFile = Join-Path $outputDir "Work_Item_Ids.txt"
            $workItemIds | Out-File -FilePath $workItemIdsFile -Encoding UTF8 -NoNewline
            Write-Host "Work Item IDs written to: $workItemIdsFile" -ForegroundColor Green
        }

        # Write raw PR metadata as JSON for downstream scripts that need the
        # unformatted fields — e.g. Jira issue key extraction reads the branch
        # name, title, and description (Get-JiraWorkItems.ps1)
        if ($pr) {
            $prMetadataFile = Join-Path $outputDir "PR_Metadata.json"
            @{
                title         = $pr.title
                description   = $pr.description
                sourceRefName = $pr.sourceRefName
                isDraft       = [bool]$pr.isDraft
            } | ConvertTo-Json -Depth 3 | Out-File -FilePath $prMetadataFile -Encoding UTF8
            Write-Host "PR metadata written to: $prMetadataFile" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to write output file: $_"
    }
}

#endregion
