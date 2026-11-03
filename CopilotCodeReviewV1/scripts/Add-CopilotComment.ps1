<#
.SYNOPSIS
    Posts a comment to a pull request in Azure DevOps.

.DESCRIPTION
    This script is used by GitHub Copilot to add a comment to a pull request.
    It simplifies the calling process by populating the necessary parameters automatically
    from environment variables set by the pipeline task. Supports both general PR-level
    comments and file-specific inline comments.

.PARAMETER Comment
    Optional. The comment text to post. Supports markdown formatting.
    Either -Comment or -CommentFile must be provided. Refused when the review task
    runs the script: use -CommentFile instead.

.PARAMETER CommentFile
    Optional. Path to a file containing the comment text. The file content is read
    and used as the comment. This avoids shell escaping issues with complex markdown.
    Either -Comment or -CommentFile must be provided. When the review task runs the
    script, the file must be inside the working directory, outside any .git folder,
    and hold at most 100,000 characters.

.PARAMETER Status
    Optional. Status for the new thread: 'Active' or 'Closed'. Default is 'Active'.

.PARAMETER FilePath
    Optional. File path for inline comment (e.g., '/src/MyProject/Program.cs').
    When provided with StartLine, creates an inline comment on the specified file.

.PARAMETER StartLine
    Optional. Starting line number for inline comment. Required when FilePath is provided.

.PARAMETER EndLine
    Optional. Ending line number for inline comment. Defaults to StartLine if not provided.

.PARAMETER ThreadId
    Optional. The ID of an existing thread to reply to. If not specified, a new thread is created.

.EXAMPLE
    .\Add-CopilotComment.ps1 -CommentFile ./_comment.md -Status 'Closed'
    Creates a new general comment thread with closed status.

.EXAMPLE
    .\Add-CopilotComment.ps1 -CommentFile ./_comment.md -Status 'Active' -FilePath '/src/Program.cs' -StartLine 42
    Creates an inline comment on line 42 using comment text read from a file.

.EXAMPLE
    .\Add-CopilotComment.ps1 -CommentFile ./_comment.md -Status 'Active' -FilePath '/src/Program.cs' -StartLine 42 -EndLine 50
    Creates an inline comment spanning lines 42-50 of the specified file.

.EXAMPLE
    .\Add-CopilotComment.ps1 -CommentFile ./_reply.md -ThreadId 123
    Replies to thread #123.

.EXAMPLE
    .\Add-CopilotComment.ps1 -Comment "This looks good!" -Status 'Closed'
    Posts inline text (only outside a review run; the review task requires -CommentFile).

.NOTES
    Author: Little Fort Software
    Date: December 2025
    Requires: PowerShell 5.1 or later
    
    Environment Variables Used:
    - AZUREDEVOPS_TOKEN: Authentication token (PAT or OAuth)
    - AZUREDEVOPS_AUTH_TYPE: 'Basic' for PAT, 'Bearer' for OAuth
    - AZUREDEVOPS_COLLECTION_URI: Azure DevOps collection URI
    - PROJECT: Azure DevOps project name
    - REPOSITORY: Repository name
    - PRID: Pull request ID
    - ITERATION_ID: (Optional) PR iteration ID for inline comments
    - COPILOT_REVIEW_STATE_DIR: (Set by the review task) folder holding the
      review policy the comment scripts enforce
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Comment text to post")]
    [string]$Comment,

    [Parameter(Mandatory = $false, HelpMessage = "Path to a file containing the comment text")]
    [string]$CommentFile,

    # Validated in the script body rather than with [ValidateSet]: a binding
    # failure would stop the script before it can record the failed attempt
    [Parameter(Mandatory = $false, HelpMessage = "Status for the new thread: Active or Closed")]
    [string]$Status = 'Active',

    [Parameter(Mandatory = $false, HelpMessage = "File path for inline comment (e.g., '/src/MyProject/Program.cs')")]
    [string]$FilePath,

    # The numbers below are strings parsed in the script body for the same
    # reason: a non-numeric value would otherwise fail during binding
    [Parameter(Mandatory = $false, HelpMessage = "Starting line number for inline comment")]
    [string]$StartLine,

    [Parameter(Mandatory = $false, HelpMessage = "Ending line number for inline comment (defaults to StartLine)")]
    [string]$EndLine,

    [Parameter(Mandatory = $false, HelpMessage = "Existing thread ID to reply to instead of creating a new thread")]
    [string]$ThreadId
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

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Posting-activity evidence for the task. Recorded before any validation so
# that every early exit below leaves a started attempt without a done marker.
$wrapperActivityId = [guid]::NewGuid().ToString('N')
Write-ReviewActivity -Id $wrapperActivityId -Phase 'started' -Action 'wrapper'

if ($Status -notin @('Active', 'Closed')) {
    Write-Error "Invalid -Status '$Status'. Valid values: Active, Closed."
    exit 1
}

# Parse the numeric arguments (0 = not provided). New variable names: assigning
# a number back to a [string] parameter would make later comparisons string-based.
$numbers = @{}
foreach ($name in 'StartLine', 'EndLine', 'ThreadId') {
    $raw = [string](Get-Variable -Name $name -ValueOnly)
    $parsed = 0
    if (-not [string]::IsNullOrWhiteSpace($raw) -and
        (-not [int]::TryParse($raw.Trim(), [ref]$parsed) -or $parsed -lt 1)) {
        Write-Error "Invalid -$name '$raw'. Provide a whole number of 1 or more."
        exit 1
    }
    $numbers[$name] = $parsed
}
$startLineNumber = $numbers.StartLine
$endLineNumber = $numbers.EndLine
$threadIdNumber = $numbers.ThreadId

# Resolve comment text from either -Comment or -CommentFile. During a review
# (PR #65 rework) the text must come from a file inside the working directory:
# inline text passes through the agent's shell command line, where command
# substitution in text copied from the pull request could run. The script this
# wrapper calls enforces the rest of the review policy.
$secrets = Get-ReviewSecretValues -RequestToken ${env:AZUREDEVOPS_TOKEN}
$policy = Get-ReviewPolicy -Secrets $secrets
if ($policy) {
    if (-not [string]::IsNullOrEmpty($Comment)) {
        Exit-ReviewPolicyRefusal -Secrets $secrets -Reason 'comment text cannot be passed inline with -Comment during a review; write it to a file in the working directory and pass -CommentFile.'
    }
    if ($CommentFile) {
        # Untyped on purpose: $Comment is a [string] parameter, which would
        # turn the helper's $null (file not found) into an empty string
        $fileText = Read-ReviewPolicyTextFile -Policy $policy -Secrets $secrets -Path $CommentFile -ParameterName '-CommentFile'
        if ($null -eq $fileText) {
            Write-Error "Comment file not found: $CommentFile"
            exit 1
        }
        $Comment = $fileText
    }
}
elseif ($CommentFile) {
    if (-not (Test-Path $CommentFile)) {
        Write-Error "Comment file not found: $CommentFile"
        exit 1
    }
    $Comment = Get-Content -Path $CommentFile -Raw
}

if (-not $Comment) {
    Write-Error "Either -Comment or -CommentFile must be provided."
    exit 1
}

# The posting script checks everything it sends; this earlier check keeps a
# credential in the text or path out of the progress lines printed below
Assert-NoReviewSecret -Secrets $secrets -Fields ([ordered]@{ 'comment text' = $Comment; 'file path' = $FilePath })

# Use the provided Status parameter (default: Active). The prompt should pass -Status when possible.
Write-Host "Posting comment with thread status: $Status" -ForegroundColor DarkGray

# Build the base parameters
$params = @{
    Token        = ${env:AZUREDEVOPS_TOKEN}
    AuthType     = ${env:AZUREDEVOPS_AUTH_TYPE}
    CollectionUri = ${env:AZUREDEVOPS_COLLECTION_URI}
    Project      = ${env:PROJECT}
    Repository   = ${env:REPOSITORY}
    Id           = ${env:PRID}
    Comment      = $Comment
    Status       = $Status
}

# Add thread ID for replies to existing threads
if ($threadIdNumber -gt 0) {
    $params.ThreadId = $threadIdNumber
    Write-Host "Replying to thread #$threadIdNumber" -ForegroundColor DarkGray
}

# Add inline comment parameters if file path is provided
if ($FilePath) {
    $params.FilePath = $FilePath

    if ($startLineNumber -gt 0) {
        $params.StartLine = $startLineNumber

        # Default EndLine to StartLine if not provided
        if ($endLineNumber -gt 0) {
            $params.EndLine = $endLineNumber
        } else {
            $params.EndLine = $startLineNumber
        }
    }
    
    # Add iteration ID if available
    if (${env:ITERATION_ID}) {
        $params.IterationId = [int]${env:ITERATION_ID}
    }
    
    Write-Host "Posting inline comment on file: $FilePath (Lines $($params.StartLine)-$($params.EndLine))" -ForegroundColor DarkGray
}

& "$scriptDir\Add-AzureDevOpsPRComment.ps1" @params

# The call operator runs the child script in-process: its 'exit' sets
# $LASTEXITCODE here without terminating this script, so failures must be
# propagated explicitly or the calling agent sees a successful exit (issue #56)
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-ReviewActivity -Id $wrapperActivityId -Phase 'done' -Action 'wrapper'
