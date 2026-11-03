<#
.SYNOPSIS
    Deletes a comment from a pull request thread in Azure DevOps.

.DESCRIPTION
    This script is used by GitHub Copilot to delete previously-created comments.
    It can delete individual comments within a thread. Note that deleting all comments
    in a thread does not delete the thread itself. The script is designed to fail
    silently to avoid blocking the code review process.

.PARAMETER ThreadId
    Required. The ID of the thread containing the comment.

.PARAMETER CommentId
    Required. The ID of the comment to delete.

.EXAMPLE
    .\Delete-CopilotComment.ps1 -ThreadId 123 -CommentId 456
    Deletes comment #456 from thread #123.

.NOTES
    Author: Little Fort Software
    Date: February 2026
    Requires: PowerShell 5.1 or later

    Environment Variables Used:
    - AZUREDEVOPS_TOKEN: Authentication token (PAT or OAuth)
    - AZUREDEVOPS_AUTH_TYPE: 'Basic' for PAT, 'Bearer' for OAuth
    - AZUREDEVOPS_COLLECTION_URI: Azure DevOps collection URI
    - PROJECT: Azure DevOps project name
    - REPOSITORY: Repository name
    - PRID: Pull request ID

    This script fails silently by design. Errors are suppressed to avoid
    interrupting the Copilot review workflow. The exception is the review
    policy: when the review task runs the script, it refuses (exit 1) to act on
    another pull request or on a comment that no automated review wrote.

    Important: You can only delete comments that were created by the same identity
    (Build Service or PAT user) making the delete request.
#>

[CmdletBinding()]
param(
    # The IDs are strings validated in the script body: a binding failure would
    # stop the script before it can record the failed attempt
    [Parameter(Mandatory = $false, HelpMessage = "Thread ID containing the comment")]
    [string]$ThreadId,

    [Parameter(Mandatory = $false, HelpMessage = "Comment ID to delete")]
    [string]$CommentId
)

function Write-ReviewActivity {
    <#
        Records review-agent posting activity for the task (issue #67): an
        "<id>.started" file when an attempt begins and "<id>.done" when it
        succeeds. The task records an iteration as reviewed only if every
        started attempt finished, so any exit before "done" counts as a failed
        attempt without instrumenting each failure path. Best-effort: never
        throws, and does nothing outside the task (variable unset).
    #>
    param(
        [string]$Id,
        [string]$Phase,
        [string]$Action
    )

    if ([string]::IsNullOrWhiteSpace($env:COPILOT_REVIEW_STATE_DIR)) {
        return
    }
    try {
        $activityDir = Join-Path $env:COPILOT_REVIEW_STATE_DIR 'activity'
        [void][System.IO.Directory]::CreateDirectory($activityDir)
        [System.IO.File]::WriteAllText((Join-Path $activityDir "$Id.$Phase"), [string]$Action)
    }
    catch {
        # Bookkeeping must never fail a comment post
    }
}

#region Review policy
# Posting controls for review runs (PR #65 rework). Before the review agent
# starts, the task writes review-policy.json to its per-run state folder
# (COPILOT_REVIEW_STATE_DIR): the pull request this run may act on, the threads
# and comments the agent may touch, the working directory, and the files the
# pull request changes. The comment scripts check every action against it.
#
# Keep this region identical in Add-CopilotComment.ps1,
# Add-AzureDevOpsPRComment.ps1, Update-CopilotComment.ps1, and
# Delete-CopilotComment.ps1 (checked by tests/powershell/ReviewPolicy.Tests.ps1).

function Get-ReviewSecretValues {
    <#
        Values that must never appear in anything the scripts send: the token
        the script authenticates with and the credentials the task and the
        agent CLIs hold. Each also counts in Base64 and in the Basic-auth form
        (Base64 of ':<value>'). Values under 8 characters are skipped as too
        likely to occur by chance. Encoded or split values still get through;
        this stops the naive case. Returns Label/Value pairs.
    #>
    param([string]$RequestToken)

    $sources = [System.Collections.Generic.List[object]]::new()
    $sources.Add([pscustomobject]@{ Label = 'the Azure DevOps token'; Value = $RequestToken })
    foreach ($name in @('AZUREDEVOPS_TOKEN', 'SYSTEM_ACCESSTOKEN', 'ANTHROPIC_API_KEY', 'CLAUDE_CODE_OAUTH_TOKEN',
            'GH_TOKEN', 'GITHUB_TOKEN', 'COPILOT_GITHUB_TOKEN', 'COPILOT_PROVIDER_API_KEY', 'COPILOT_PROVIDER_BEARER_TOKEN')) {
        $sources.Add([pscustomobject]@{ Label = $name; Value = [System.Environment]::GetEnvironmentVariable($name) })
    }

    $secrets = [System.Collections.Generic.List[object]]::new()
    foreach ($source in $sources) {
        $value = ([string]$source.Value).Trim()
        if ($value.Length -lt 8) {
            continue
        }
        $secrets.Add([pscustomobject]@{ Label = $source.Label; Value = $value })
        foreach ($form in @($value, ":$value")) {
            $encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($form))
            $secrets.Add([pscustomobject]@{ Label = "$($source.Label) (Base64)"; Value = $encoded })
        }
    }
    return $secrets.ToArray()
}

function Hide-ReviewSecret {
    # Replaces every secret value in a message with '***'
    param(
        [string]$Text,
        $Secrets
    )

    foreach ($secret in $Secrets) {
        $Text = $Text.Replace([string]$secret.Value, '***')
    }
    return $Text
}

function Exit-ReviewPolicyRefusal {
    <#
        Refuses the requested action and exits 1. The reason goes to the caller
        (the review agent) on stderr and to the task, which reports it as a
        pipeline warning, with secret values masked. The caller's 'started'
        activity marker never gets its 'done' marker, so the run is not
        recorded as having reviewed the iteration.
    #>
    param(
        [string]$Reason,
        $Secrets
    )

    $message = Hide-ReviewSecret -Text $Reason -Secrets $Secrets
    [Console]::Error.WriteLine("Refused by review policy: $message")
    if (-not [string]::IsNullOrWhiteSpace($env:COPILOT_REVIEW_STATE_DIR)) {
        try {
            $activityDir = Join-Path $env:COPILOT_REVIEW_STATE_DIR 'activity'
            [void][System.IO.Directory]::CreateDirectory($activityDir)
            $refusalFile = Join-Path $activityDir "$([guid]::NewGuid().ToString('N')).refused"
            [System.IO.File]::WriteAllText($refusalFile, $message)
        }
        catch {
            # Reporting is best-effort; the refusal itself still happens
        }
    }
    exit 1
}

function Get-ReviewPolicy {
    <#
        The review policy for this run, or $null outside the task. Inside the
        task a missing or unreadable policy refuses every action, so the checks
        can never silently switch off.
    #>
    param($Secrets)

    if ([string]::IsNullOrWhiteSpace($env:COPILOT_REVIEW_STATE_DIR)) {
        return $null
    }
    $policyFile = Join-Path $env:COPILOT_REVIEW_STATE_DIR 'review-policy.json'
    $policy = $null
    try {
        $policy = [System.IO.File]::ReadAllText($policyFile) | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $policy = $null
    }
    # Every field the checks read must be present with the right type: a
    # missing completeness flag would otherwise read as 'incomplete' and
    # loosen the inline-path check
    $valid = $policy -is [System.Management.Automation.PSCustomObject] -and $policy.schemaVersion -eq 1 -and
        $policy.destination -and $policy.destination.collectionUri -and $policy.destination.project -and
        $policy.destination.repository -and $policy.destination.pullRequestId -and $policy.workspaceRoot -and
        $policy.threadsComplete -is [bool] -and $policy.threads -is [array] -and
        $policy.changedFilesComplete -is [bool] -and $policy.changedFiles -is [array]
    if (-not $valid) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the review policy for this run ($policyFile) is missing or unreadable, so no action is allowed."
    }
    return $policy
}

function ConvertTo-ReviewPolicyCollectionUri {
    # Comparable form of a collection URI (scheme, host, port, and path in lower
    # case, no trailing slash), or $null for a URI with credentials, a query,
    # or a fragment, which never matches
    param([string]$Uri)

    $parsed = $null
    if (-not [System.Uri]::TryCreate(([string]$Uri).Trim(), [System.UriKind]::Absolute, [ref]$parsed) -or
        $parsed.UserInfo -or $parsed.Query -or $parsed.Fragment) {
        return $null
    }
    return ('{0}://{1}:{2}{3}' -f $parsed.Scheme, $parsed.IdnHost, $parsed.Port, $parsed.AbsolutePath.TrimEnd('/')).ToLowerInvariant()
}

function Assert-ReviewPolicyDestination {
    <#
        Refuses unless the script is about to call the pull request the task
        was configured for. Runs before any request: a misled agent could
        otherwise point the script at another server and send it the token.
    #>
    param(
        $Policy,
        $Secrets,
        [string]$CollectionUri,
        [string]$Project,
        [string]$Repository,
        [string]$PullRequestId
    )

    $expected = $Policy.destination
    $expectedUri = ConvertTo-ReviewPolicyCollectionUri ([string]$expected.collectionUri)
    $expectedId = 0
    $requestedId = 0
    $same = $expectedUri -and $expectedUri -eq (ConvertTo-ReviewPolicyCollectionUri $CollectionUri) -and
        [string]::Equals([string]$expected.project, $Project, [System.StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals([string]$expected.repository, $Repository, [System.StringComparison]::OrdinalIgnoreCase) -and
        [int]::TryParse([string]$expected.pullRequestId, [ref]$expectedId) -and
        [int]::TryParse(([string]$PullRequestId).Trim(), [ref]$requestedId) -and
        $expectedId -eq $requestedId
    if (-not $same) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason ("this review may act only on pull request #{0} in {1}/{2} at {3}, not on pull request #{4} in {5}/{6} at {7}." -f
            $expected.pullRequestId, $expected.project, $expected.repository, $expected.collectionUri,
            $PullRequestId, $Project, $Repository, $CollectionUri)
    }
}

function Get-ReviewPolicyThread {
    # The policy entry for a thread, or $null when the thread is not listed
    param(
        $Policy,
        [long]$ThreadId
    )

    foreach ($thread in @($Policy.threads)) {
        if ($null -ne $thread -and [long]$thread.threadId -eq $ThreadId) {
            return $thread
        }
    }
    return $null
}

function Assert-ReviewPolicyThread {
    <#
        Refuses unless the policy lets this review reply to the thread or change
        its status. The task allows threads from earlier automated reviews and,
        with resolveThreads 'all', active and pending human threads; raw prompts
        may act on any thread.
    #>
    param(
        $Policy,
        $Secrets,
        [long]$ThreadId,
        [ValidateSet('Reply', 'UpdateStatus')]
        [string]$Action
    )

    $verb = if ($Action -eq 'Reply') { 'reply to' } else { 'change the status of' }
    if ($Policy.threadsComplete -ne $true) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the pull request's comment threads could not be loaded for this review, so it may not $verb thread #$ThreadId."
    }
    $thread = Get-ReviewPolicyThread -Policy $Policy -ThreadId $ThreadId
    $permission = if ($null -eq $thread) { $false } elseif ($Action -eq 'Reply') { $thread.mayReply } else { $thread.mayUpdateStatus }
    if ($permission -ne $true) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "this review may not $verb thread #$ThreadId; it may act only on the threads listed in the JSON sections of PR_Details.txt."
    }
}

function Assert-ReviewPolicyAgentComment {
    # Refuses unless the comment was written by an automated review (for edits
    # and deletions)
    param(
        $Policy,
        $Secrets,
        [long]$ThreadId,
        [long]$CommentId,
        [string]$Verb
    )

    if ($Policy.threadsComplete -ne $true) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the pull request's comment threads could not be loaded for this review, so it may not $Verb comment #$CommentId in thread #$ThreadId."
    }
    $thread = Get-ReviewPolicyThread -Policy $Policy -ThreadId $ThreadId
    $isAgentComment = $false
    if ($null -ne $thread) {
        foreach ($agentCommentId in @($thread.agentCommentIds)) {
            if ($null -ne $agentCommentId -and [long]$agentCommentId -eq $CommentId) {
                $isAgentComment = $true
            }
        }
    }
    if (-not $isAgentComment) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "comment #$CommentId in thread #$ThreadId was not written by an automated review, so this review may not $Verb it."
    }
}

function Resolve-ReviewPolicyPath {
    <#
        Absolute path with every symbolic link and junction along it resolved,
        so a link inside the working directory cannot lead to a file elsewhere.
        Relative paths start at the current PowerShell location. Parts that do
        not exist are kept as they are. $null when the path cannot be resolved
        (for example a link loop).
    #>
    param([string]$Path)

    try {
        $separators = [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
        $full = [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path))
        # Each pass resolves the first link on the path and starts over, since
        # the link's target can contain links of its own
        for ($pass = 0; $pass -lt 40; $pass++) {
            $root = [System.IO.Path]::GetPathRoot($full)
            $parts = $full.Substring($root.Length).Split($separators, [System.StringSplitOptions]::RemoveEmptyEntries)
            $current = $root
            $linkFound = $false
            for ($i = 0; $i -lt $parts.Count; $i++) {
                # GetFullPath removes these, except in \\?\ paths
                if ($parts[$i] -eq '.' -or $parts[$i] -eq '..') {
                    return $null
                }
                $current = [System.IO.Path]::Combine($current, $parts[$i])
                $item = if ([System.IO.Directory]::Exists($current)) { [System.IO.DirectoryInfo]::new($current) } else { [System.IO.FileInfo]::new($current) }
                $target = $item.LinkTarget
                if ($target) {
                    # A relative link target starts at the link's own folder
                    if (-not [System.IO.Path]::IsPathRooted($target)) {
                        $target = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($current), $target)
                    }
                    $remaining = @($parts | Select-Object -Skip ($i + 1))
                    $full = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine([string[]](@($target) + $remaining)))
                    $linkFound = $true
                    break
                }
            }
            if (-not $linkFound) {
                return $full
            }
        }
    }
    catch {
        # Unresolvable paths are refused by the callers
    }
    return $null
}

function Get-ReviewPolicyPathProblem {
    # Why a resolved path may not be used as a comment source: outside the
    # resolved working directory, or inside a .git folder (which can hold
    # credentials). $null when the path is fine.
    param(
        [string]$Path,
        [string]$Folder
    )

    $separators = [char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $comparison = if ($IsWindows) { [System.StringComparison]::OrdinalIgnoreCase } else { [System.StringComparison]::Ordinal }
    $prefix = $Folder.TrimEnd($separators) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $Path.StartsWith($prefix, $comparison)) {
        return 'is outside the working directory'
    }
    foreach ($part in $Path.Substring($prefix.Length).Split($separators, [System.StringSplitOptions]::RemoveEmptyEntries)) {
        if ([string]::Equals($part, '.git', [System.StringComparison]::OrdinalIgnoreCase)) {
            return 'is inside a .git folder'
        }
    }
    return $null
}

function Read-ReviewPolicyTextFile {
    <#
        Reads comment text the review agent wrote to a file. During a review the
        file must lie inside the working directory and outside any .git folder
        (after resolving links) and hold at most 100,000 characters, so the
        scripts cannot be used to post the contents of other files such as
        credentials or CLI configuration. Returns $null when the file does not
        exist (the caller reports that); other problems are refusals.
    #>
    param(
        $Policy,
        $Secrets,
        [string]$Path,
        [string]$ParameterName
    )

    $maxCharacters = 100000
    $folder = Resolve-ReviewPolicyPath ([string]$Policy.workspaceRoot)
    $file = Resolve-ReviewPolicyPath $Path
    $problem = if (-not $folder -or -not $file) { 'could not be resolved' } else { Get-ReviewPolicyPathProblem -Path $file -Folder $folder }
    if ($problem) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the $ParameterName file '$Path' $problem; write the text to a file inside the working directory ($($Policy.workspaceRoot))."
    }
    if (-not [System.IO.File]::Exists($file)) {
        return $null
    }
    # At most 4 bytes per character, so a larger file is certainly too long
    if ([System.IO.FileInfo]::new($file).Length -gt 4 * $maxCharacters) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the $ParameterName file '$Path' is larger than a comment may be (100,000 characters)."
    }
    $text = [System.IO.File]::ReadAllText($file)
    if ($text.Length -gt $maxCharacters) {
        Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the $ParameterName file '$Path' has $($text.Length) characters; a comment may have at most 100,000."
    }
    return $text
}

function Test-ReviewPolicyInlinePath {
    <#
        Whether an inline comment may anchor to a file: one of the files the
        pull request changes (exact match on the Azure DevOps path, such as
        '/src/App.cs') or, when that list is incomplete, a file in the working
        directory.
    #>
    param(
        $Policy,
        [string]$NormalizedPath
    )

    if ($Policy.changedFilesComplete -eq $true) {
        foreach ($changedFile in @($Policy.changedFiles)) {
            if ([string]::Equals([string]$changedFile, $NormalizedPath, [System.StringComparison]::Ordinal)) {
                return $true
            }
        }
        return $false
    }
    $folder = Resolve-ReviewPolicyPath ([string]$Policy.workspaceRoot)
    if (-not $folder) {
        return $false
    }
    $file = Resolve-ReviewPolicyPath ([System.IO.Path]::Combine($folder, $NormalizedPath.TrimStart('/')))
    return [bool]($file -and -not (Get-ReviewPolicyPathProblem -Path $file -Folder $folder) -and [System.IO.File]::Exists($file))
}

function Assert-NoReviewSecret {
    # Refuses when any outgoing text contains a secret value (see
    # Get-ReviewSecretValues). Fields: ordered name => text.
    param(
        $Secrets,
        $Fields
    )

    foreach ($field in $Fields.GetEnumerator()) {
        $text = [string]$field.Value
        if (-not $text) {
            continue
        }
        foreach ($secret in $Secrets) {
            if ($text.Contains([string]$secret.Value)) {
                Exit-ReviewPolicyRefusal -Secrets $Secrets -Reason "the $($field.Key) contains the value of $($secret.Label); credentials must never be posted."
            }
        }
    }
}

#endregion Review policy

# Posting-activity evidence for the task. Every exit before the 'done' marker
# below, including the non-blocking failure paths, counts as a failed attempt.
# A completed deletion is not counted as a posted comment.
$activityId = [guid]::NewGuid().ToString('N')
Write-ReviewActivity -Id $activityId -Phase 'started' -Action 'delete'

# Wrap entire script in try/catch for silent failure
try {
    # Parse the IDs (new variable names: assigning a number back to a [string]
    # parameter would make later comparisons string-based)
    $threadIdNumber = 0
    $commentIdNumber = 0
    if (-not [int]::TryParse(([string]$ThreadId).Trim(), [ref]$threadIdNumber) -or $threadIdNumber -lt 1 -or
        -not [int]::TryParse(([string]$CommentId).Trim(), [ref]$commentIdNumber) -or $commentIdNumber -lt 1) {
        Write-Warning "Delete-CopilotComment: -ThreadId and -CommentId must both be whole numbers of 1 or more (got '$ThreadId' and '$CommentId')."
        exit 0
    }

    # Read credentials from environment variables
    $token = ${env:AZUREDEVOPS_TOKEN}
    $authType = ${env:AZUREDEVOPS_AUTH_TYPE}
    $collectionUri = ${env:AZUREDEVOPS_COLLECTION_URI}
    $project = ${env:PROJECT}
    $repository = ${env:REPOSITORY}
    $prId = ${env:PRID}

    # Validate required environment variables
    if ([string]::IsNullOrEmpty($token) -or
        [string]::IsNullOrEmpty($collectionUri) -or
        [string]::IsNullOrEmpty($project) -or
        [string]::IsNullOrEmpty($repository) -or
        [string]::IsNullOrEmpty($prId)) {
        # Missing required env vars - log a warning so the issue is visible in pipeline logs
        $missing = @()
        if ([string]::IsNullOrEmpty($token)) { $missing += 'AZUREDEVOPS_TOKEN' }
        if ([string]::IsNullOrEmpty($collectionUri)) { $missing += 'AZUREDEVOPS_COLLECTION_URI' }
        if ([string]::IsNullOrEmpty($project)) { $missing += 'PROJECT' }
        if ([string]::IsNullOrEmpty($repository)) { $missing += 'REPOSITORY' }
        if ([string]::IsNullOrEmpty($prId)) { $missing += 'PRID' }
        Write-Warning "Delete-CopilotComment: Skipping deletion of comment #$commentIdNumber in thread #$threadIdNumber — required environment variable(s) not set: $($missing -join ', ')"
        exit 0
    }

    # Review policy (PR #65 rework): during a review, only comments written by
    # an automated review, and only on the configured pull request. Refusals
    # exit 1 (the catch below never sees them).
    $secrets = Get-ReviewSecretValues -RequestToken $token
    $policy = Get-ReviewPolicy -Secrets $secrets
    if ($policy) {
        Assert-ReviewPolicyDestination -Policy $policy -Secrets $secrets -CollectionUri $collectionUri `
            -Project $project -Repository $repository -PullRequestId $prId
        Assert-ReviewPolicyAgentComment -Policy $policy -Secrets $secrets -ThreadId $threadIdNumber `
            -CommentId $commentIdNumber -Verb 'delete'
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

    # Build the API URL for deleting a comment
    $baseUrl = "$collectionUri/$project/_apis"
    $uri = "$baseUrl/git/repositories/$repository/pullrequests/$prId/threads/$threadIdNumber/comments/$commentIdNumber`?api-version=7.1"

    # Send DELETE request
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Delete -ErrorAction Stop

    # Log success (visible in pipeline logs but doesn't affect workflow)
    Write-Host "Comment #$commentIdNumber in thread #$threadIdNumber deleted" -ForegroundColor Green

    Write-ReviewActivity -Id $activityId -Phase 'done' -Action 'delete'
}
catch {
    # Non-blocking failure - log detailed error info but don't fail the pipeline
    $errorMsg = "Delete-CopilotComment: Could not delete comment #$CommentId in thread #$ThreadId"

    $statusCode = $null
    $errorDetail = $null
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
    }
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        $errorDetail = $_.ErrorDetails.Message
    }

    if ($statusCode) {
        $errorMsg += " (HTTP $statusCode)"
    }
    if ($errorDetail) {
        $errorMsg += " — API response: $errorDetail"
    }
    elseif ($_.Exception.Message) {
        $errorMsg += " — $($_.Exception.Message)"
    }

    Write-Warning $errorMsg
}

# Always exit with success
exit 0
