<#
.SYNOPSIS
    Retrieves details for Jira issues referenced by a pull request.

.DESCRIPTION
    This script extracts Jira issue keys (e.g. PROJ-123) from a pull request's
    source branch name, title, and description, fetches each issue from the
    Jira Cloud REST API, and writes the details to a structured text file for
    use as context in code reviews.

    Candidate keys are self-validating: anything matching the key pattern is
    tried against the API, and 404 responses are silently skipped. This filters
    out false positives such as 'UTF-8' without needing a precise regex.

.PARAMETER BaseUrl
    Required. The Jira Cloud site base URL (e.g., 'https://yourorg.atlassian.net').

.PARAMETER Email
    Required. The Atlassian account email associated with the API token.

.PARAMETER ApiToken
    Required. A Jira Cloud API token created at https://id.atlassian.com/manage-profile/security/api-tokens.

.PARAMETER PrMetadataFile
    Required. Path to the PR_Metadata.json file written by Get-AzureDevOpsPR.ps1,
    containing the pull request's title, description, and source branch name.

.PARAMETER ProjectKeys
    Optional. Comma-separated list of Jira project keys (e.g., 'PROJ,CORE').
    When provided, only issue keys with these prefixes are considered.

.PARAMETER MaxIssues
    Optional. Maximum number of candidate issue keys to fetch. Default is 10.

.PARAMETER OutputFile
    Optional. Path to write the output to a file. If not specified, output is only written to the console.

.EXAMPLE
    .\Get-JiraWorkItems.ps1 -BaseUrl "https://myorg.atlassian.net" -Email "me@example.com" -ApiToken "token" -PrMetadataFile "./PR_Metadata.json" -OutputFile "./Work_Item_Details.txt"
    Extracts issue keys from the PR metadata and writes the fetched issue details to a file.

.NOTES
    Author: Little Fort Software
    Date: August 2026
    Requires: PowerShell 7 or later

    Jira Cloud only — authentication uses Basic auth with email:apiToken, which
    Jira Server/Data Center does not support (those use Bearer PATs).

    Exit codes: 0 on success (including no issues found), 1 on configuration or
    authentication failures.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Jira Cloud site base URL (e.g., https://yourorg.atlassian.net)")]
    [ValidateNotNullOrEmpty()]
    [string]$BaseUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Atlassian account email for API token authentication")]
    [ValidateNotNullOrEmpty()]
    [string]$Email,

    [Parameter(Mandatory = $true, HelpMessage = "Jira Cloud API token")]
    [ValidateNotNullOrEmpty()]
    [string]$ApiToken,

    [Parameter(Mandatory = $true, HelpMessage = "Path to the PR_Metadata.json file with title, description, and source branch")]
    [ValidateNotNullOrEmpty()]
    [string]$PrMetadataFile,

    [Parameter(Mandatory = $false, HelpMessage = "Comma-separated Jira project keys to restrict extraction (e.g., 'PROJ,CORE')")]
    [string]$ProjectKeys,

    [Parameter(Mandatory = $false, HelpMessage = "Maximum number of candidate issue keys to fetch")]
    [ValidateRange(1, 50)]
    [int]$MaxIssues = 10,

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

function Get-JiraIssueKeys {
    param(
        [string]$Text,
        [string]$ProjectKeys,
        [int]$MaxKeys = 10
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    # Jira key convention: project key (letter followed by letters/digits) plus
    # a numeric suffix. Matched case-insensitively because branch names commonly
    # lowercase the key; canonical Jira keys are uppercase. False positives
    # (e.g. 'UTF-8') are filtered later by the API 404 check.
    $found = [regex]::Matches($Text, '\b([A-Za-z][A-Za-z0-9]+)-(\d+)\b') |
        ForEach-Object { $_.Value.ToUpperInvariant() }

    # Deduplicate while preserving discovery order
    $unique = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $found) {
        if (-not $unique.Contains($key)) {
            $unique.Add($key) | Out-Null
        }
    }

    # Restrict to configured project keys if provided
    if (-not [string]::IsNullOrWhiteSpace($ProjectKeys)) {
        $allowedPrefixes = ($ProjectKeys -split ',') | ForEach-Object { $_.Trim().ToUpperInvariant() } | Where-Object { $_ }
        $unique = @($unique | Where-Object {
            $keyPrefix = ($_ -split '-')[0]
            $allowedPrefixes -contains $keyPrefix
        })
    }

    return @($unique | Select-Object -First $MaxKeys)
}

function Invoke-JiraApi {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    # Returns the response object, or $null for 404 (candidate key does not
    # exist — expected for false-positive matches). All other failures throw
    # after printing diagnostics, since they affect every request (e.g. auth).
    try {
        return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -ErrorAction Stop
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        if ($statusCode -eq 404) {
            return $null
        }

        Write-Host "##[error]Jira API call failed: GET $Uri"
        if ($statusCode) {
            Write-Host "##[error]HTTP status: $statusCode"
        }
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            Write-Host "##[error]API response: $($_.ErrorDetails.Message)"
        }
        else {
            Write-Host "##[error]$($_.Exception.Message)"
        }
        if ($statusCode -eq 401) {
            Write-Host "##[error]Authentication failed. Verify the Jira email and API token (created at https://id.atlassian.com/manage-profile/security/api-tokens)."
        }
        elseif ($statusCode -eq 403) {
            Write-Host "##[error]Permission denied. Verify the account has browse access to the Jira project."
        }
        throw
    }
}

function ConvertFrom-Html {
    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    # Convert strikethrough tags to ~~text~~ BEFORE stripping tags, so retracted
    # content (e.g. superseded acceptance criteria) stays marked as struck-through
    # instead of reading as live requirements (issue #52). The \b prevents 's'
    # from matching the start of unrelated tags like <span> or <strong>.
    $text = $Html -replace '(?is)<(del|strike|s)\b[^>]*>(.*?)</\1\s*>', '~~$2~~'
    # Convert <br> tags to newlines
    $text = $text -replace '<br\s*/?>', "`n"
    # Convert block-level tags to newlines
    $text = $text -replace '</(p|div|li|tr|h[1-6])>', "`n"
    $text = $text -replace '<(p|div|li|tr|h[1-6])[^>]*>', ""
    # Strip all remaining HTML tags
    $text = $text -replace '<[^>]+>', ''
    # Decode HTML entities
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    # Collapse multiple consecutive blank lines into one
    $text = $text -replace "(\r?\n\s*){3,}", "`n`n"
    # Trim leading/trailing whitespace
    $text = $text.Trim()

    return $text
}

#endregion

#region Main Logic

# Initialize output handling
$script:OutputToFile = -not [string]::IsNullOrEmpty($OutputFile)
$script:OutputBuilder = [System.Text.StringBuilder]::new()

# Normalize the base URL: strip trailing slashes
$BaseUrl = $BaseUrl.TrimEnd('/')

# Read PR metadata written by Get-AzureDevOpsPR.ps1
if (-not (Test-Path $PrMetadataFile)) {
    Write-Host "##[error]PR metadata file not found: $PrMetadataFile"
    exit 1
}

try {
    $prMetadata = Get-Content -Path $PrMetadataFile -Raw | ConvertFrom-Json
}
catch {
    Write-Host "##[error]Failed to parse PR metadata file: $_"
    exit 1
}

# Scan the branch name, title, and description for candidate issue keys
$searchText = @($prMetadata.sourceRefName, $prMetadata.title, $prMetadata.description) -join "`n"
$candidateKeys = @(Get-JiraIssueKeys -Text $searchText -ProjectKeys $ProjectKeys -MaxKeys $MaxIssues)

if ($candidateKeys.Count -eq 0) {
    Write-Host "No Jira issue keys found in the PR branch name, title, or description." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nCandidate Jira issue key(s): $($candidateKeys -join ', ')" -ForegroundColor Cyan

# Basic auth: base64(email:apiToken) — Jira Cloud convention
$base64Auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${Email}:${ApiToken}"))
$headers = @{
    Authorization = "Basic $base64Auth"
    Accept        = "application/json"
}

# Fetch each candidate; renderedFields returns HTML for rich-text fields.
# 404s are skipped (false-positive keys); any other API failure aborts with
# exit 1 — diagnostics were already printed by Invoke-JiraApi.
$fields = "summary,issuetype,status,priority,description"
$issues = @()
try {
    foreach ($key in $candidateKeys) {
        $issueUrl = "$BaseUrl/rest/api/3/issue/$key`?fields=$fields&expand=renderedFields"
        $issue = Invoke-JiraApi -Uri $issueUrl -Headers $headers

        if ($null -eq $issue) {
            Write-Host "  $key — not found in Jira (skipping; likely a false-positive match)" -ForegroundColor DarkGray
            continue
        }

        $issues += $issue
        Write-Host "  $key — found: $($issue.fields.summary)" -ForegroundColor Green
    }
}
catch {
    exit 1
}

if ($issues.Count -eq 0) {
    Write-Host "None of the candidate keys resolved to Jira issues." -ForegroundColor Yellow
    exit 0
}

# Display results
Write-Output-Line ("=" * 80) -ForegroundColor DarkGray
Write-Output-Line "LINKED WORK ITEM DETAILS (JIRA)" -ForegroundColor Green
Write-Output-Line ("=" * 80) -ForegroundColor DarkGray

foreach ($issue in $issues) {
    $issueType = $issue.fields.issuetype.name
    $summary = $issue.fields.summary
    $status = $issue.fields.status.name
    $priority = if ($issue.fields.priority) { $issue.fields.priority.name } else { $null }
    # renderedFields carries the description as HTML rendered from Jira's
    # native format; fall back to nothing if rendering is unavailable
    $description = ConvertFrom-Html $issue.renderedFields.description

    Write-Output-Line "`n[$($issue.key) - $issueType]" -ForegroundColor Yellow
    Write-Output-Line "  Title:           $summary"
    Write-Output-Line "  Status:          $status"
    if ($priority) {
        Write-Output-Line "  Priority:        $priority"
    }

    if (-not [string]::IsNullOrWhiteSpace($description)) {
        Write-Output-Line "`n  Description:"
        $description -split "`n" | ForEach-Object {
            Write-Output-Line "    $_"
        }
    }
}

Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray

# Write to output file if specified
if ($script:OutputToFile) {
    try {
        $outputDir = Split-Path -Parent $OutputFile
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        $script:OutputBuilder.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "`nOutput written to: $OutputFile" -ForegroundColor Green
    }
    catch {
        Write-Host "##[error]Failed to write output file: $_"
        exit 1
    }
}

exit 0

#endregion
