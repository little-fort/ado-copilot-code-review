<#
.SYNOPSIS
    Retrieves commits and changed files from the most recent iteration of a pull request.

.DESCRIPTION
    This script uses the Azure DevOps REST API to get the list of commits and 
    changed files from the most recent iteration (latest push) of a pull request.

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
    Required. The pull request ID to retrieve changes for.

.PARAMETER ReviewStateKeyPrefix
    Optional (pass together with -ReviewStateFile). Enables incremental re-reviews (issue #67): the
    pull request properties named <prefix>.Iteration.<n> record which iterations this review
    configuration has already reviewed (written by Set-AzureDevOpsPRReviewState.ps1). When a usable
    record exists, a [Changes Since Last Review] section is added to the output.

.PARAMETER ReviewStateFile
    Optional (pass together with -ReviewStateKeyPrefix). Path of the JSON file that receives the
    current iteration, the last reviewed iteration, and the changes since then, for the task to act on.

.PARAMETER UntrustedContentId
    Optional. Random per-run id (32 lowercase hex characters) from the task. When set, the output file
    wraps the commit list and the changed-file lists (PR-author text) in untrusted-content markers.

.EXAMPLE
    .\Get-AzureDevOpsPRChanges.ps1 -Token "your-pat" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123
    Retrieves the commits and changed files from the most recent iteration of PR #123.

.EXAMPLE
    .\Get-AzureDevOpsPRChanges.ps1 -Token "oauth-token" -AuthType "Bearer" -CollectionUri "https://dev.azure.com/myorg" -Project "myproject" -Repository "myrepo" -Id 123
    Retrieves PR changes using OAuth/System.AccessToken authentication.

.EXAMPLE
    .\Get-AzureDevOpsPRChanges.ps1 -Token "your-pat" -CollectionUri "https://tfs.contoso.com/tfs/DefaultCollection" -Project "myproject" -Repository "myrepo" -Id 123 -OutputFile "C:\output\pr-changes.txt"
    Writes the pull request changes to the specified file (on-prem example).

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

    [Parameter(Mandatory = $true, HelpMessage = "Repository name")]
    [ValidateNotNullOrEmpty()]
    [string]$Repository,

    [Parameter(Mandatory = $true, HelpMessage = "Pull request ID")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Id,

    [Parameter(Mandatory = $false, HelpMessage = "Output file path to write results to")]
    [string]$OutputFile,

    [Parameter(Mandatory = $false, HelpMessage = "Pull request property prefix identifying the review configuration")]
    [ValidatePattern('^[A-Za-z0-9._-]{1,200}$')]
    [string]$ReviewStateKeyPrefix,

    [Parameter(Mandatory = $false, HelpMessage = "Path of the review state JSON file to write")]
    [string]$ReviewStateFile,

    [Parameter(Mandatory = $false, HelpMessage = "Random per-run id for untrusted-content markers")]
    [ValidatePattern('^[0-9a-f]{32}$')]
    [string]$UntrustedContentId
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

function Write-UntrustedContentMarker {
    <#
        Writes the opening or closing line of an untrusted-content block to the
        output file (PR #65 rework). Text between the markers was written by PR
        authors, commenters, or work item authors; the review prompt tells the
        agent to treat it as data, never as instructions. The id is random per
        run (passed by the task), so the content cannot forge a closing marker.
        Writes nothing without an id or an output file. The markers go to the
        context file only; the console output stays as it was.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Begin', 'End')]
        [string]$Kind,

        [string]$Source = ''
    )

    if ([string]::IsNullOrEmpty($script:UntrustedContentId) -or -not $script:OutputToFile) {
        return
    }
    $marker = if ($Kind -eq 'Begin') {
        "<<<UNTRUSTED_CONTENT id=$($script:UntrustedContentId) source=$Source>>>"
    }
    else {
        "<<<END_UNTRUSTED_CONTENT id=$($script:UntrustedContentId)>>>"
    }
    $script:OutputBuilder.AppendLine($marker) | Out-Null
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
        [string]$Method = "Get",
        # Optional: receives the response headers (needed for endpoints that
        # page via a continuation-token header).
        [ref]$ResponseHeaders
    )

    try {
        $params = @{
            Uri         = $Uri
            Headers     = $Headers
            Method      = $Method
            ErrorAction = 'Stop'
        }
        $captureHeaders = $ResponseHeaders -and $PSVersionTable.PSVersion.Major -ge 6
        if ($captureHeaders) {
            $params.ResponseHeadersVariable = 'capturedHeaders'
        }

        $response = Invoke-RestMethod @params
        if ($captureHeaders) {
            $ResponseHeaders.Value = $capturedHeaders
        }
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
        $baseMsg += " calling $Method $Uri"

        if ($statusCode -eq 401) {
            Write-Error "$baseMsg — Authentication failed. Please verify your token is valid and has appropriate permissions. API response: $errorDetail"
        }
        elseif ($statusCode -eq 404) {
            Write-Error "$baseMsg — Resource not found. Please verify the organization, project, repository, and PR ID. API response: $errorDetail"
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

function Format-DateForDisplay {
    param([string]$DateString)
    
    if ([string]::IsNullOrEmpty($DateString)) {
        return "N/A"
    }
    
    try {
        $date = [DateTime]::Parse($DateString)
        return $date.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return $DateString
    }
}

function Get-ChangeTypeDisplay {
    param([string]$ChangeType)
    
    switch ($ChangeType) {
        "add"      { return @{ Text = "Added"; Color = "Green" } }
        "edit"     { return @{ Text = "Modified"; Color = "Yellow" } }
        "delete"   { return @{ Text = "Deleted"; Color = "Red" } }
        "rename"   { return @{ Text = "Renamed"; Color = "Cyan" } }
        "copy"     { return @{ Text = "Copied"; Color = "Cyan" } }
        default    { return @{ Text = $ChangeType; Color = "White" } }
    }
}

function Get-ContinuationToken {
    # Returns the x-ms-continuationtoken response header value, or $null when
    # absent. Header names are matched case-insensitively.
    param($ResponseHeaders)

    if (-not $ResponseHeaders) {
        return $null
    }
    foreach ($name in $ResponseHeaders.Keys) {
        if ($name -ieq 'x-ms-continuationtoken') {
            $value = @($ResponseHeaders[$name])[0]
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return [string]$value
            }
        }
    }
    return $null
}

function Get-PullRequestCommits {
    <#
        Retrieves every commit in the pull request (issue #64). The commits
        endpoint pages via a continuation token: when more commits exist, the
        response carries an x-ms-continuationtoken header to send back as the
        continuationToken query parameter.

        Returns Items plus Status: 'Complete', 'Partial', or 'Failed'.
    #>
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$ApiVersion,
        [int]$MaxPages = 50
    )

    $commits = [System.Collections.Generic.List[object]]::new()
    $seenTokens = [System.Collections.Generic.HashSet[string]]::new()
    $token = $null

    for ($pageNumber = 1; $pageNumber -le $MaxPages; $pageNumber++) {
        $uri = "$BaseUrl/commits?$ApiVersion"
        if ($token) {
            $uri += "&continuationToken=$([uri]::EscapeDataString($token))"
        }

        $responseHeaders = $null
        $page = Invoke-AzureDevOpsApi -Uri $uri -Headers $Headers -ResponseHeaders ([ref]$responseHeaders)
        if ($null -eq $page) {
            $status = if ($pageNumber -eq 1) { 'Failed' } else { 'Partial' }
            return [pscustomobject]@{ Items = $commits; Status = $status }
        }
        if ($page.value) {
            $commits.AddRange([object[]]@($page.value))
        }

        $token = Get-ContinuationToken -ResponseHeaders $responseHeaders
        if (-not $token) {
            return [pscustomobject]@{ Items = $commits; Status = 'Complete' }
        }
        if (-not $seenTokens.Add($token)) {
            # A repeated token would loop forever; keep what was retrieved
            return [pscustomobject]@{ Items = $commits; Status = 'Partial' }
        }
    }

    return [pscustomobject]@{ Items = $commits; Status = 'Partial' }
}

function Get-IterationChanges {
    <#
        Retrieves every changed file in an iteration (see issue #64). The changes
        endpoint returns 100 entries per page by default (2000 at most) and
        reports the next page through nextSkip/nextTop, which are 0 on the
        last page.

        Returns Items plus Status: 'Complete', 'Partial', or 'Failed' (see
        Get-PullRequestCommits).
    #>
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$ApiVersion,
        [int]$IterationId,
        # When above 0, return only the changes since this earlier iteration
        # ($compareTo; the default compares against the merge base)
        [int]$CompareTo = 0,
        [int]$MaxPages = 100
    )

    $entries = [System.Collections.Generic.List[object]]::new()
    $skip = 0
    $top = 0

    for ($pageNumber = 1; $pageNumber -le $MaxPages; $pageNumber++) {
        $uri = "$BaseUrl/iterations/$IterationId/changes?$ApiVersion"
        if ($CompareTo -gt 0) {
            $uri += "&`$compareTo=$CompareTo"
        }
        if ($pageNumber -gt 1) {
            $uri += "&`$top=$top&`$skip=$skip"
        }

        $page = Invoke-AzureDevOpsApi -Uri $uri -Headers $Headers
        if ($null -eq $page) {
            $status = if ($pageNumber -eq 1) { 'Failed' } else { 'Partial' }
            return [pscustomobject]@{ Items = $entries; Status = $status }
        }
        if ($page.changeEntries) {
            $entries.AddRange([object[]]@($page.changeEntries))
        }

        # Missing paging fields cast to 0, which ends the loop
        $nextSkip = [int]$page.nextSkip
        $nextTop = [int]$page.nextTop
        if ($nextTop -le 0) {
            return [pscustomobject]@{ Items = $entries; Status = 'Complete' }
        }
        if ($nextSkip -le $skip) {
            # A next page that does not advance would loop forever
            return [pscustomobject]@{ Items = $entries; Status = 'Partial' }
        }
        $skip = $nextSkip
        $top = $nextTop
    }

    return [pscustomobject]@{ Items = $entries; Status = 'Partial' }
}

function Get-PullRequestProperties {
    <#
        Reads the pull request's property bag. The review state it holds is
        optional, so a failure prints a warning and the review simply falls back
        to a full review (Invoke-AzureDevOpsApi's Write-Error would paint a red
        error for a non-fatal condition). 404 means there are no properties.
        Returns Status ('Ok' or 'Failed'), the property Names, and the raw Value.
    #>
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$ApiVersion
    )

    try {
        $response = Invoke-RestMethod -Uri "$BaseUrl/properties?$ApiVersion" -Headers $Headers -Method Get -ErrorAction Stop
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        if ($statusCode -eq 404) {
            return [pscustomobject]@{ Status = 'Ok'; Names = @(); Value = $null }
        }
        $reason = if ($statusCode) { "HTTP $statusCode" } else { ConvertTo-SafeLogText $_.Exception.Message }
        Write-Host "##[warning]Could not read the incremental review state from the pull request properties ($reason); the whole pull request will be reviewed."
        return [pscustomobject]@{ Status = 'Failed'; Names = @(); Value = $null }
    }

    $names = @()
    if ($response -and $response.value) {
        $names = @($response.value.PSObject.Properties | ForEach-Object { $_.Name })
    }
    return [pscustomobject]@{ Status = 'Ok'; Names = $names; Value = $response.value }
}

function Get-PropertyValue {
    # Returns the value of one property from the bag, or $null. Each entry is
    # an object with "$type" and "$value" members (single quotes matter: in
    # double quotes "$value" would expand a variable).
    param(
        $Properties,
        [string]$Name
    )

    if ($null -eq $Properties) {
        return $null
    }
    $entry = $Properties.PSObject.Properties[$Name]
    if (-not $entry -or $null -eq $entry.Value) {
        return $null
    }
    $valueMember = $entry.Value.PSObject.Properties['$value']
    if ($valueMember) {
        return $valueMember.Value
    }
    return $entry.Value
}

function Get-ReviewState {
    <#
        Works out what this review configuration has already reviewed (issue #67).

        Each reviewed iteration is recorded as its own property,
        <prefix>.Iteration.<n> = <source commit SHA>. The last reviewed iteration
        is the highest n up to the current one whose recorded SHA still matches
        that iteration's source commit; anything else is ignored.

        When that iteration is older than the current one, the changes since then
        are fetched ($compareTo) — but only when the result is trustworthy:
          - same merge base (commonRefCommit) in both iterations, so a rebase or a
            merge from the target branch cannot pass target-branch changes off as
            the pull request's own;
          - no retarget in between (the diff base changed);
          - the full change list is complete (needed to filter the result).
        The result is narrowed to files that are still part of the pull request.
        In every other case the delta is omitted and the whole PR is reviewed.
    #>
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$ApiVersion,
        [string]$KeyPrefix,
        [object[]]$Iterations,
        $CurrentIteration,
        $FullChanges
    )

    $currentId = [int]$CurrentIteration.id
    $state = [ordered]@{
        stateReadStatus          = 'NotFound'
        lastReviewedIteration    = $null
        lastReviewedSourceCommit = $null
        delta                    = [ordered]@{ status = 'NotApplicable'; reason = $null; files = @() }
    }

    $properties = Get-PullRequestProperties -BaseUrl $BaseUrl -Headers $Headers -ApiVersion $ApiVersion
    if ($properties.Status -ne 'Ok') {
        $state.stateReadStatus = 'Failed'
        $state.delta.reason = 'The review state could not be read.'
        return $state
    }

    $iterationsById = @{}
    foreach ($iteration in $Iterations) {
        $iterationsById[[int]$iteration.id] = $iteration
    }

    $keyPattern = '^' + [regex]::Escape($KeyPrefix) + '\.Iteration\.(\d{1,9})$'
    $lastId = 0
    $lastSha = $null
    foreach ($name in $properties.Names) {
        $match = [regex]::Match($name, $keyPattern)
        if (-not $match.Success) {
            continue
        }
        $recordedId = [int]$match.Groups[1].Value
        if ($recordedId -lt 1 -or $recordedId -gt $currentId -or $recordedId -le $lastId) {
            continue
        }
        $recordedSha = [string](Get-PropertyValue -Properties $properties.Value -Name $name)
        $iteration = $iterationsById[$recordedId]
        if (-not $iteration -or -not $iteration.sourceRefCommit -or
            [string]::IsNullOrWhiteSpace($recordedSha) -or
            $recordedSha.Trim() -ine [string]$iteration.sourceRefCommit.commitId) {
            continue
        }
        $lastId = $recordedId
        $lastSha = [string]$iteration.sourceRefCommit.commitId
    }

    if ($lastId -eq 0) {
        return $state
    }

    $state.stateReadStatus = 'Found'
    $state.lastReviewedIteration = $lastId
    $state.lastReviewedSourceCommit = $lastSha

    if ($lastId -eq $currentId) {
        $state.delta.reason = 'The current iteration was already reviewed.'
        return $state
    }

    $lastBase = [string]$iterationsById[$lastId].commonRefCommit.commitId
    $currentBase = [string]$CurrentIteration.commonRefCommit.commitId
    if ([string]::IsNullOrWhiteSpace($lastBase) -or $lastBase -ine $currentBase) {
        $state.delta.reason = 'The merge base changed since the last review (for example a rebase or a merge from the target branch).'
        return $state
    }

    $retargeted = @($Iterations | Where-Object {
        [int]$_.id -gt $lastId -and [int]$_.id -le $currentId -and $_.reason -eq 'retarget'
    })
    if ($retargeted.Count -gt 0) {
        $state.delta.reason = 'The pull request was retargeted since the last review.'
        return $state
    }

    if ($FullChanges.Status -ne 'Complete') {
        $state.delta.reason = 'The full change list is incomplete.'
        return $state
    }

    Write-Host "Retrieving changes since reviewed iteration #$lastId..." -ForegroundColor Cyan
    $deltaResult = Get-IterationChanges -BaseUrl $BaseUrl -Headers $Headers -ApiVersion $ApiVersion `
        -IterationId $currentId -CompareTo $lastId
    if ($deltaResult.Status -ne 'Complete') {
        $state.delta.status = $deltaResult.Status
        $state.delta.reason = 'The changes since the last review could not be fully retrieved.'
        return $state
    }

    # Keep only files still in the pull request's changes (drops files the
    # pull request later reverted to their target-branch content)
    $prPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($change in $FullChanges.Items) {
        if ($change.item -and $change.item.path) {
            [void]$prPaths.Add([string]$change.item.path)
        }
    }
    $state.delta.status = 'Complete'
    $state.delta.files = @($deltaResult.Items | Where-Object {
        $_.item -and $_.item.path -and $prPaths.Contains([string]$_.item.path)
    } | ForEach-Object {
        [ordered]@{
            path         = [string]$_.item.path
            changeType   = [string]$_.changeType
            originalPath = $_.originalPath
        }
    })
    return $state
}

#endregion

#region Main Logic

# Initialize output handling
$script:OutputToFile = -not [string]::IsNullOrEmpty($OutputFile)
$script:OutputBuilder = [System.Text.StringBuilder]::new()

if ([string]::IsNullOrEmpty($ReviewStateKeyPrefix) -ne [string]::IsNullOrEmpty($ReviewStateFile)) {
    Write-Error "-ReviewStateKeyPrefix and -ReviewStateFile must be passed together."
    exit 1
}

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$baseUrl = "$CollectionUri/$Project/_apis/git/repositories/$Repository/pullrequests/$Id"
$apiVersion = "api-version=7.1"

# Verify the PR exists
Write-Host "`nRetrieving pull request #$Id..." -ForegroundColor Cyan
$prUrl = "$baseUrl`?$apiVersion"
$pr = Invoke-AzureDevOpsApi -Uri $prUrl -Headers $headers

if ($null -eq $pr) {
    Write-Error "Failed to retrieve pull request #$Id from repository '$Repository'. See the error above for details."
    exit 1
}

Write-Host "Found PR: $(ConvertTo-SafeLogText $pr.title)" -ForegroundColor Green
Write-Host "Status: $($pr.status.ToUpper())" -ForegroundColor $(if ($pr.status -eq "active") { "Green" } else { "Yellow" })

# Get iterations
Write-Host "`nRetrieving iterations..." -ForegroundColor Cyan
$iterationsUrl = "$baseUrl/iterations?$apiVersion"
$iterations = Invoke-AzureDevOpsApi -Uri $iterationsUrl -Headers $headers

if ($null -eq $iterations -or $iterations.count -eq 0) {
    Write-Warning "No iterations found for this pull request."
    exit 0
}

$latestIteration = $iterations.value | Sort-Object -Property id -Descending | Select-Object -First 1
$iterationId = $latestIteration.id

Write-Host "Found $($iterations.count) iteration(s). Using latest: Iteration #$iterationId" -ForegroundColor Green

# Get all commits for the PR (paged)
Write-Host "`nRetrieving commits..." -ForegroundColor Cyan
$commitResult = Get-PullRequestCommits -BaseUrl $baseUrl -Headers $headers -ApiVersion $apiVersion
if ($commitResult.Status -eq 'Partial') {
    Write-Host "##[warning]Only $($commitResult.Items.Count) commit(s) could be retrieved; the commit list for this PR is incomplete."
}

# Get all changed files for the latest iteration (paged)
Write-Host "Retrieving changes for iteration #$iterationId..." -ForegroundColor Cyan
$changeResult = Get-IterationChanges -BaseUrl $baseUrl -Headers $headers -ApiVersion $apiVersion -IterationId $iterationId
if ($changeResult.Status -eq 'Partial') {
    Write-Host "##[warning]Only $($changeResult.Items.Count) changed file(s) could be retrieved; the changed-file list for this PR is incomplete."
}
$changeEntries = $changeResult.Items

# Incremental re-review state (issue #67), only when the task asks for it
$reviewState = $null
if ($ReviewStateKeyPrefix) {
    Write-Host "Reading incremental review state..." -ForegroundColor Cyan
    $reviewState = Get-ReviewState -BaseUrl $baseUrl -Headers $headers -ApiVersion $apiVersion `
        -KeyPrefix $ReviewStateKeyPrefix -Iterations @($iterations.value) `
        -CurrentIteration $latestIteration -FullChanges $changeResult
    if ($reviewState.lastReviewedIteration) {
        Write-Host "Last reviewed iteration: #$($reviewState.lastReviewedIteration)" -ForegroundColor Green
    }
    if ($reviewState.delta.reason) {
        Write-Host "  $($reviewState.delta.reason)" -ForegroundColor DarkGray
    }
}

# Display results
Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray
Write-Output-Line "PULL REQUEST CHANGES - ITERATION #$iterationId" -ForegroundColor Green
Write-Output-Line ("=" * 80) -ForegroundColor DarkGray

# Iteration Info
Write-Output-Line "`n[Iteration Details]" -ForegroundColor Yellow
Write-Output-Line "  Iteration ID:     #$iterationId"
Write-Output-Line "  Created:          $(Format-DateForDisplay $latestIteration.createdDate)"
Write-Output-Line "  Updated:          $(Format-DateForDisplay $latestIteration.updatedDate)"
# Full SHAs with old/new labels so the review agent diffs in the correct direction
# (see the diff-direction guidance in prompt.txt)
if ($latestIteration.sourceRefCommit) {
    Write-Output-Line "  Source Commit (PR branch, new code):   $($latestIteration.sourceRefCommit.commitId)"
}
if ($latestIteration.targetRefCommit) {
    Write-Output-Line "  Target Commit (base branch, old code): $($latestIteration.targetRefCommit.commitId)"
}

# Commits
Write-Output-Line "`n[Commits in this PR]" -ForegroundColor Yellow
if ($commitResult.Items.Count -gt 0) {
    if ($commitResult.Status -eq 'Partial') {
        Write-Output-Line "  Commits retrieved: $($commitResult.Items.Count) (INCOMPLETE: Azure DevOps did not return the full list; more commits exist)`n"
    }
    else {
        Write-Output-Line "  Total commits: $($commitResult.Items.Count)`n"
    }

    # Commit messages and author names are PR-author text
    Write-UntrustedContentMarker -Kind Begin -Source 'commits'
    foreach ($commit in $commitResult.Items) {
        $shortId = $commit.commitId.Substring(0, 8)
        $message = $commit.comment -split "`n" | Select-Object -First 1
        if ($message.Length -gt 60) {
            $message = $message.Substring(0, 57) + "..."
        }
        Write-Output-Line "  $shortId - $message" -ForegroundColor Cyan
        Write-Output-Line "           Author: $($commit.author.name) | $(Format-DateForDisplay $commit.author.date)" -ForegroundColor DarkGray
    }
    Write-UntrustedContentMarker -Kind End
}
elseif ($commitResult.Status -eq 'Failed') {
    Write-Output-Line "  Commits could not be retrieved from Azure DevOps."
}
else {
    Write-Output-Line "  No commits found."
}

# Changed Files
Write-Output-Line "`n[Changed Files]" -ForegroundColor Yellow
if ($changeEntries.Count -gt 0) {
    # Group by change type for summary
    $addedCount = ($changeEntries | Where-Object { $_.changeType -eq "add" }).Count
    $modifiedCount = ($changeEntries | Where-Object { $_.changeType -eq "edit" }).Count
    $deletedCount = ($changeEntries | Where-Object { $_.changeType -eq "delete" }).Count
    $otherCount = $changeEntries.Count - $addedCount - $modifiedCount - $deletedCount

    if ($changeResult.Status -eq 'Partial') {
        # Tell the review agent the list is not authoritative so it falls back
        # to git for the full set of changes
        Write-Output-Line "  Files retrieved: $($changeEntries.Count) (INCOMPLETE: Azure DevOps did not return the full list; use git to identify every changed file)"
    }
    else {
        Write-Output-Line "  Total files changed: $($changeEntries.Count)"
    }
    $summaryLine = "  +$addedCount added | ~$modifiedCount modified | -$deletedCount deleted"
    if ($otherCount -gt 0) {
        $summaryLine += " | $otherCount other"
    }
    Write-Output-Line $summaryLine
    Write-Output-Line ""
    
    # List each file. File names (which may even contain line breaks) and
    # rename origins are chosen by the PR author.
    Write-UntrustedContentMarker -Kind Begin -Source 'changed-files'
    foreach ($change in $changeEntries) {
        $changeDisplay = Get-ChangeTypeDisplay -ChangeType $change.changeType
        $filePath = $change.item.path

        Write-Output-Line "  [$($changeDisplay.Text)] $filePath" -ForegroundColor $changeDisplay.Color

        # Show original path for renames
        if ($change.changeType -eq "rename" -and $change.originalPath) {
            Write-Output-Line "         (from: $($change.originalPath))" -ForegroundColor DarkGray
        }
    }
    Write-UntrustedContentMarker -Kind End
}
elseif ($changeResult.Status -eq 'Failed') {
    # Previously this case printed "No file changes found", which told the
    # review agent there was nothing to review
    Write-Output-Line "  Changed files could not be retrieved from Azure DevOps; use git to identify the changed files."
}
else {
    Write-Output-Line "  No file changes found in this iteration."
}

# Changes since the last review by this configuration. The section name is
# referenced by the prompt templates' re-review guidance (keep in lockstep).
if ($reviewState -and $reviewState.delta.status -eq 'Complete') {
    $lastId = $reviewState.lastReviewedIteration
    $lastSha = $reviewState.lastReviewedSourceCommit
    $currentSha = [string]$latestIteration.sourceRefCommit.commitId
    $deltaFiles = @($reviewState.delta.files)

    Write-Output-Line "`n[Changes Since Last Review]" -ForegroundColor Yellow
    Write-Output-Line "  Last reviewed iteration: #$lastId (source commit $lastSha)"
    if ($deltaFiles.Count -eq 0) {
        Write-Output-Line "  No files in this pull request changed since that review."
    }
    else {
        Write-Output-Line "  Files in this pull request changed since that review: $($deltaFiles.Count)"
        Write-UntrustedContentMarker -Kind Begin -Source 'changed-files-since-last-review'
        foreach ($file in $deltaFiles) {
            $changeDisplay = Get-ChangeTypeDisplay -ChangeType $file.changeType
            Write-Output-Line "    [$($changeDisplay.Text)] $($file.path)" -ForegroundColor $changeDisplay.Color
            if ($file.originalPath) {
                Write-Output-Line "           (from: $($file.originalPath))" -ForegroundColor DarkGray
            }
        }
        Write-UntrustedContentMarker -Kind End
        Write-Output-Line "  To see exactly what changed since that review: git diff $lastSha $currentSha"
        Write-Output-Line "  If that commit is not available locally (for example after a force push or with a shallow clone), use this file list together with the full pull request diff."
    }
}

Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray

# Provide link to the PR
$webUrl = "$CollectionUri/$Project/_git/$Repository/pullrequest/$Id"
Write-Host "`nView PR: $webUrl" -ForegroundColor Cyan
if ($script:OutputToFile) {
    $script:OutputBuilder.AppendLine("`nView PR: $webUrl") | Out-Null
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
        
        # Also write the iteration ID to a separate file for use by other scripts
        $iterationIdFile = Join-Path $outputDir "Iteration_Id.txt"
        $iterationId.ToString() | Out-File -FilePath $iterationIdFile -Encoding UTF8 -NoNewline
        Write-Host "Iteration ID written to: $iterationIdFile" -ForegroundColor Green

        # Write full commit SHAs for use by diffOnlyReview mode
        if ($latestIteration.sourceRefCommit) {
            $sourceCommitFile = Join-Path $outputDir "Source_Commit.txt"
            $latestIteration.sourceRefCommit.commitId | Out-File -FilePath $sourceCommitFile -Encoding UTF8 -NoNewline
            Write-Host "Source commit SHA written to: $sourceCommitFile" -ForegroundColor Green
        }
        if ($latestIteration.targetRefCommit) {
            $targetCommitFile = Join-Path $outputDir "Target_Commit.txt"
            $latestIteration.targetRefCommit.commitId | Out-File -FilePath $targetCommitFile -Encoding UTF8 -NoNewline
            Write-Host "Target commit SHA written to: $targetCommitFile" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to write output file: $_"
    }
}

# Review state for the task: what was fetched now plus what was already
# reviewed. Written whenever requested, even without a usable record, so the
# task always knows the current iteration and commit this run captured.
if ($ReviewStateFile) {
    $stateDocument = [ordered]@{
        schemaVersion            = 1
        stateKeyPrefix           = $ReviewStateKeyPrefix
        currentIteration         = [int]$iterationId
        currentSourceCommit      = [string]$latestIteration.sourceRefCommit.commitId
        fullChangesStatus        = [string]$changeResult.Status
        stateReadStatus          = $reviewState.stateReadStatus
        lastReviewedIteration    = $reviewState.lastReviewedIteration
        lastReviewedSourceCommit = $reviewState.lastReviewedSourceCommit
        delta                    = $reviewState.delta
    }
    try {
        $stateDir = Split-Path -Parent $ReviewStateFile
        if (-not [string]::IsNullOrEmpty($stateDir) -and -not (Test-Path $stateDir)) {
            New-Item -ItemType Directory -Path $stateDir -Force | Out-Null
        }
        $stateJson = ConvertTo-Json -InputObject $stateDocument -Depth 6
        [System.IO.File]::WriteAllText($ReviewStateFile, $stateJson, [System.Text.UTF8Encoding]::new($false))
        Write-Host "Review state written to: $ReviewStateFile" -ForegroundColor Green
    }
    catch {
        # The review itself does not depend on this file; without it the task
        # neither skips nor records this run
        Write-Host "##[warning]Could not write the review state file: $(ConvertTo-SafeLogText "$_")"
    }
}

#endregion
