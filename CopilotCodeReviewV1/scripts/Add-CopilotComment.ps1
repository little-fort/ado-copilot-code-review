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
    Either -Comment or -CommentFile must be provided.

.PARAMETER CommentFile
    Optional. Path to a file containing the comment text. The file content is read
    and used as the comment. This avoids shell escaping issues with complex markdown.
    Either -Comment or -CommentFile must be provided.

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
    .\Add-CopilotComment.ps1 -Comment "This looks good!" -Status 'Closed'
    Creates a new general comment thread with closed status.

.EXAMPLE
    .\Add-CopilotComment.ps1 -CommentFile ./_comment.md -Status 'Active' -FilePath '/src/Program.cs' -StartLine 42
    Creates an inline comment on line 42 using comment text read from a file.

.EXAMPLE
    .\Add-CopilotComment.ps1 -Comment "This block needs refactoring" -Status 'Active' -FilePath '/src/Program.cs' -StartLine 42 -EndLine 50
    Creates an inline comment spanning lines 42-50 of the specified file.

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

# Resolve comment text from either -Comment or -CommentFile
if ($CommentFile) {
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
