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
            $errorDetail = $_.ErrorDetails.Message
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
            Write-Error "$baseMsg — $($_.Exception.Message)"
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
    
    Write-Output-Line "`n[Comments/Threads]" -ForegroundColor Yellow
    if ($threads -and $threads.value) {
        # Filter to top-level comment threads (exclude system-generated threads)
        $commentThreads = $threads.value | Where-Object { 
            $_.comments -and 
            $_.comments.Count -gt 0 -and 
            $_.comments[0].commentType -ne "system"
        }
        
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
                
                # Display comment content (truncate if too long)
                $commentContent = $firstComment.content
                if (-not [string]::IsNullOrEmpty($commentContent)) {
                    # Clean up and format comment
                    $commentLines = $commentContent -split "`n"
                    $displayLines = $commentLines | Select-Object -First 30
                    foreach ($line in $displayLines) {
                        $trimmedLine = $line.Trim()
                        if (-not [string]::IsNullOrEmpty($trimmedLine)) {
                            Write-Output-Line "    $trimmedLine"
                        }
                    }
                    if ($commentLines.Count -gt 3) {
                        Write-Output-Line "    ... ($($commentLines.Count - 3) more lines)" -ForegroundColor DarkGray
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

    # Output structured JSON for Copilot-created comment threads
    # This section enables Copilot to identify and resolve its previous comments
    if ($threads -and $threads.value) {
        $commentThreads = $threads.value | Where-Object { 
            $_.comments -and 
            $_.comments.Count -gt 0 -and 
            $_.comments[0].commentType -ne "system"
        }
        
        # Filter to threads that were created by Copilot
        # Match by: Build Service author OR comment contains [Generated by GitHub Copilot]
        $copilotThreads = $commentThreads | Where-Object {
            # Check if author is Build Service identity
            ($_.comments[0].author.displayName -like "*Build Service*") -or
            ($_.comments[0].author.uniqueName -like "*Build Service*") -or
            # Check if comment contains Copilot tag
            ($_.comments[0].content -like "*[Generated by GitHub Copilot]*")
        }
        
        if ($copilotThreads -and $copilotThreads.Count -gt 0) {
            Write-Output-Line "`n=== COPILOT COMMENT THREADS (JSON) ===" -ForegroundColor Cyan
            
            $jsonThreads = $copilotThreads | ForEach-Object {
                $thread = $_
                $firstComment = $thread.comments | Select-Object -First 1

                # Extract file path and line info
                $filePath = $null
                $startLine = $null
                if ($thread.threadContext -and $thread.threadContext.filePath) {
                    $filePath = $thread.threadContext.filePath
                    if ($thread.threadContext.rightFileStart) {
                        $startLine = $thread.threadContext.rightFileStart.line
                    }
                    elseif ($thread.threadContext.leftFileStart) {
                        $startLine = $thread.threadContext.leftFileStart.line
                    }
                }

                # Build replies array from all comments after the first
                $replies = @()
                $allComments = @($thread.comments)
                if ($allComments.Count -gt 1) {
                    $replies = $allComments | Select-Object -Skip 1 | Where-Object {
                        $_.commentType -ne "system"
                    } | ForEach-Object {
                        @{
                            author        = $_.author.displayName
                            content       = $_.content
                            publishedDate = $_.publishedDate
                        }
                    }
                    if ($null -eq $replies) { $replies = @() }
                }

                @{
                    threadId  = $thread.id
                    commentId = $firstComment.id
                    status    = $thread.status
                    filePath  = $filePath
                    startLine = $startLine
                    content   = $firstComment.content
                    replies   = @($replies)
                }
            }

            $jsonOutput = $jsonThreads | ConvertTo-Json -Compress -Depth 5
            Write-Output-Line $jsonOutput
            Write-Output-Line "=== END COPILOT COMMENT THREADS ===" -ForegroundColor Cyan
        }
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
    Write-Host $tableOutput
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
    }
    catch {
        Write-Error "Failed to write output file: $_"
    }
}

#endregion
