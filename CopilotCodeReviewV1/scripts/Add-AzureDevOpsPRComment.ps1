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

    Review policy: when the review task runs the script (COPILOT_REVIEW_STATE_DIR set),
    it refuses (exit 1, before any request) to post to another pull request or server
    and to reply to threads the review may not act on. Inline comments on a file the
    pull request does not change are posted as general comments with the file and
    line appended. Outside the task, only the secret check applies: comment text,
    file paths, and request bodies may never contain the token or another known
    credential.

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

function ConvertTo-SafeLogText {
    <#
        Neutralizes Azure Pipelines logging-command markers in text that came
        from outside this script (API responses, PR and work item content)
        before it is written to the pipeline log (issues #47, #69).
    #>
    param([string]$Text)

    return $Text -replace '##(?=(?:vso)?\[)', '# #'
}

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
                $statusText = if ($statusCode) { "HTTP $statusCode" } else { ConvertTo-SafeLogText $_.Exception.Message }
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
            # External text (response bodies can echo request content) goes
            # through ConvertTo-SafeLogText; the ##[error] prefix stays intact
            if ($responseBody) {
                Write-Host "##[error]API response: $(ConvertTo-SafeLogText $responseBody)"
            }
            else {
                Write-Host "##[error]$(ConvertTo-SafeLogText $_.Exception.Message)"
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

# Posting-activity evidence for the task (see Write-ReviewActivity)
$activityId = [guid]::NewGuid().ToString('N')
Write-ReviewActivity -Id $activityId -Phase 'started' -Action $(if ($ThreadId -gt 0) { 'reply' } else { 'thread' })

# Review policy (PR #65 rework; see the region above). Every check runs before
# the first request, so a refused action sends nothing anywhere.
$secrets = Get-ReviewSecretValues -RequestToken $Token
$policy = Get-ReviewPolicy -Secrets $secrets
if ($policy) {
    Assert-ReviewPolicyDestination -Policy $policy -Secrets $secrets -CollectionUri $CollectionUri `
        -Project $Project -Repository $Repository -PullRequestId ([string]$Id)
    if ($ThreadId -gt 0) {
        Assert-ReviewPolicyThread -Policy $policy -Secrets $secrets -ThreadId $ThreadId -Action Reply
    }
}

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$baseUrl = "$CollectionUri/$Project/_apis/git/repositories/$Repository/pullrequests/$Id"
$apiVersion = "api-version=7.1"
$threadsUrl = "$baseUrl/threads?$apiVersion"

# Inline anchoring applies to new threads only; replies ignore the file parameters
$isInlineComment = $ThreadId -le 0 -and -not [string]::IsNullOrEmpty($FilePath) -and $StartLine -gt 0
$normalizedPath = if ($FilePath) { Format-AzureDevOpsFilePath -Path $FilePath } else { '' }
$effectiveEndLine = if ($EndLine -gt 0) { $EndLine } else { $StartLine }

# General-comment form of an inline comment, with the file and line in the text
$lineInfo = if ($StartLine -eq $effectiveEndLine) { "Line $StartLine" } else { "Lines $StartLine-$effectiveEndLine" }
$fallbackComment = $Comment + "`n`n**File:** ``$normalizedPath```n**$lineInfo**"

# Secret check, part 1 (always on, also outside the task): the agent's text and
# file path, before anything derived from them is printed
Assert-NoReviewSecret -Secrets $secrets -Fields ([ordered]@{
    'comment text'         = $Comment
    'file path'            = $FilePath
    'normalized file path' = $normalizedPath
})

# During a review, inline comments anchor only to files the pull request
# changes. A comment on any other file (say, an unchanged file the change
# affects) is posted as a general comment that names the file and line.
$anchorInline = $isInlineComment
if ($policy -and $isInlineComment -and -not (Test-ReviewPolicyInlinePath -Policy $policy -NormalizedPath $normalizedPath)) {
    $anchorInline = $false
    Write-Host "Note: $(ConvertTo-SafeLogText $normalizedPath) is not a file changed in this pull request, so the comment is posted as a general comment that names the file and line." -ForegroundColor Yellow
}

# Request bodies, built up front so the second part of the secret check
# covers exactly what could be sent
$replyBody = @{
    content         = $Comment
    parentCommentId = 0
    commentType     = 1  # Text comment
}
$postedContent = if ($isInlineComment -and -not $anchorInline) { $fallbackComment } else { $Comment }
$threadBody = @{
    comments = @(
        @{
            content     = $postedContent
            commentType = 1  # Text comment
        }
    )
    status   = Get-ThreadStatusValue -StatusName $Status
}
if ($anchorInline) {
    $threadBody.threadContext = @{
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
        $threadBody.pullRequestThreadContext = @{
            iterationContext = @{
                firstComparingIteration  = $IterationId
                secondComparingIteration = $IterationId
            }
        }
    }
}
# Used if Azure DevOps rejects the inline anchor
$fallbackBody = @{
    comments = @(
        @{
            content     = $fallbackComment
            commentType = 1
        }
    )
    status   = Get-ThreadStatusValue -StatusName $Status
}

# Secret check, part 2: every request body this run could send (the fallback
# body joins the comment and the path into one text)
$outgoing = [ordered]@{}
if ($ThreadId -gt 0) {
    $outgoing['request body'] = $replyBody | ConvertTo-Json -Depth 10
}
else {
    $outgoing['request body'] = $threadBody | ConvertTo-Json -Depth 10
    if ($anchorInline) {
        $outgoing['fallback request body'] = $fallbackBody | ConvertTo-Json -Depth 10
    }
}
Assert-NoReviewSecret -Secrets $secrets -Fields $outgoing

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

Write-Host "Found PR: $(ConvertTo-SafeLogText $pr.title)" -ForegroundColor Green

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

    # Duplicate guard for ambiguous POST failures: re-read the thread and treat
    # a reply with the exact same text as the committed result
    $verifyReplyCommitted = {
        $thread = Invoke-AzureDevOpsApi -Uri $threadUrl -Headers $headers
        @($thread.comments) | Select-Object -Skip 1 | Where-Object { $_.content -eq $Comment } | Select-Object -First 1
    }

    try {
        $result = Invoke-AzureDevOpsApi -Uri $commentsUrl -Headers $headers -Method "Post" -Body $replyBody -VerifyCompleted $verifyReplyCommitted
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
        Write-Host "  Author:       $(ConvertTo-SafeLogText $result.author.displayName)"
        Write-Host "  Posted:       $($result.publishedDate)"
        Write-Host "`n  Content:"
        # Review comments often quote code, which may itself contain logging
        # commands (e.g. when reviewing pipeline YAML)
        Write-Host "  $(ConvertTo-SafeLogText $Comment)" -ForegroundColor White
        Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
    }
}
else {
    # Create new thread
    if ($anchorInline) {
        Write-Host "`nCreating inline comment thread on $(ConvertTo-SafeLogText $normalizedPath) (Lines $StartLine-$effectiveEndLine)..." -ForegroundColor Cyan
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
            $_.comments -and @($_.comments)[0].content -eq $postedContent
        } | Select-Object -First 1
    }

    # Attempt to post the comment
    try {
        $result = Invoke-AzureDevOpsApi -Uri $threadsUrl -Headers $headers -Method "Post" -Body $threadBody -VerifyCompleted $verifyThreadCommitted
    }
    catch {
        if ($anchorInline) {
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
        if ($anchorInline -and -not $inlineCommentFailed) {
            Write-Host "  Type:         Inline comment"
            Write-Host "  File:         $(ConvertTo-SafeLogText $normalizedPath)"
            Write-Host "  Lines:        $StartLine-$effectiveEndLine"
        } else {
            Write-Host "  Type:         General comment"
        }
        Write-Host "  Comment ID:   #$($result.comments[0].id)"
        Write-Host "  Author:       $(ConvertTo-SafeLogText $result.comments[0].author.displayName)"
        Write-Host "  Posted:       $($result.comments[0].publishedDate)"
        Write-Host "`n  Content:"
        Write-Host "  $(ConvertTo-SafeLogText $Comment)" -ForegroundColor White
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

Write-ReviewActivity -Id $activityId -Phase 'done' -Action $(if ($ThreadId -gt 0) { 'reply' } else { 'thread' })

# Explicit success exit so callers reading $LASTEXITCODE never see a stale value
exit 0

#endregion
