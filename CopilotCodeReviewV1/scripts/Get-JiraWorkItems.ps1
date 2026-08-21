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
    Both classic and scoped API tokens are supported: classic tokens authenticate
    against the site URL directly, while scoped tokens (the only kind service
    accounts can create) are routed through the Atlassian platform gateway
    automatically and must include the 'read:jira-work' and 'read:jira-user' scopes.

.PARAMETER PrMetadataFile
    Required. Path to the PR_Metadata.json file written by Get-AzureDevOpsPR.ps1,
    containing the pull request's title, description, and source branch name.

.PARAMETER ProjectKeys
    Optional. Comma-separated list of Jira project keys (e.g., 'PROJ,CORE').
    When provided, only issue keys with these prefixes are considered.

.PARAMETER CustomFields
    Optional. Comma-separated list of Jira custom fields to include in the
    output, given as display names (case-insensitive, e.g. 'Acceptance
    Criteria') or field IDs (e.g. 'customfield_10031'). Display names are
    resolved through GET /rest/api/3/field; unresolvable names are skipped
    with a warning. Names containing commas cannot be expressed — use the
    field ID instead.

.PARAMETER MaxIssues
    Optional. Maximum number of candidate issue keys to fetch. Default is 10.

.PARAMETER GatewayBaseUrl
    Optional. Base URL of the Atlassian platform gateway used for scoped API
    tokens. Defaults to https://api.atlassian.com; overridable only so tests
    can point it at a mock API. Not exposed as a task input.

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

    Atlassian's two API token types are byte-indistinguishable, so the script
    probes GET /rest/api/3/myself to find the working API root: the site URL
    for classic tokens, or the platform gateway (api.atlassian.com/ex/jira/
    {cloudId}) for scoped tokens. The cloud ID is discovered unauthenticated
    via {site}/_edge/tenant_info.

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

    [Parameter(Mandatory = $false, HelpMessage = "Comma-separated Jira custom fields to include (display names or customfield_NNNNN IDs)")]
    [string]$CustomFields,

    [Parameter(Mandatory = $false, HelpMessage = "Maximum number of candidate issue keys to fetch")]
    [ValidateRange(1, 50)]
    [int]$MaxIssues = 10,

    [Parameter(Mandatory = $false, HelpMessage = "Atlassian platform gateway base URL (test seam; defaults to the public gateway)")]
    [ValidateNotNullOrEmpty()]
    [string]$GatewayBaseUrl = 'https://api.atlassian.com',

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

function Test-JiraAuth {
    param(
        [string]$Root,
        [hashtable]$Headers
    )

    # Quiet auth probe. GET /myself is the only reliable auth oracle: issue
    # endpoints silently fall back to anonymous access when Basic auth fails
    # and return 404 for anything anonymous cannot see (verified against Jira
    # Cloud), so a bad token looks like "issue not found" everywhere else.
    # A 401 here is EXPECTED while probing for the right API root, which is
    # why this does not reuse Invoke-JiraApi (it prints ##[error] diagnostics).
    try {
        $me = Invoke-RestMethod -Uri "$Root/rest/api/3/myself" -Headers $Headers -Method Get -ErrorAction Stop
        return @{ Ok = $true; StatusCode = 200; DisplayName = $me.displayName }
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        return @{ Ok = $false; StatusCode = $statusCode; DisplayName = $null }
    }
}

function Get-JiraCloudId {
    param(
        [string]$BaseUrl
    )

    # Every Jira Cloud site exposes its cloud ID unauthenticated. The ID is
    # needed to address the same site through the api.atlassian.com gateway,
    # which is the only endpoint scoped API tokens can authenticate against.
    try {
        $info = Invoke-RestMethod -Uri "$BaseUrl/_edge/tenant_info" -Method Get -ErrorAction Stop
        if ($info -and $info.cloudId) {
            return [string]$info.cloudId
        }
        return $null
    }
    catch {
        return $null
    }
}

function Resolve-JiraApiRoot {
    param(
        [string]$BaseUrl,
        [string]$GatewayBaseUrl,
        [hashtable]$Headers
    )

    # Atlassian issues two kinds of API tokens that are indistinguishable by
    # inspection: classic tokens authenticate against the site URL, while
    # scoped tokens (the only kind service accounts can create) only work
    # through the platform gateway. Probe the site URL first so existing
    # classic-token configs behave exactly as before, and fall back to the
    # gateway only on HTTP 401 — the verified scoped-token signature. Any
    # other site failure (403/5xx/network) is a real problem, not a token-
    # type mismatch, so it fails without a fallback attempt.
    $siteAuth = Test-JiraAuth -Root $BaseUrl -Headers $Headers
    if ($siteAuth.Ok) {
        Write-Host "Authenticated to Jira as: $($siteAuth.DisplayName)" -ForegroundColor Green
        return $BaseUrl
    }

    if ($siteAuth.StatusCode -ne 401) {
        Write-Host "##[error]Jira authentication preflight failed: GET $BaseUrl/rest/api/3/myself"
        if ($siteAuth.StatusCode) {
            Write-Host "##[error]HTTP status: $($siteAuth.StatusCode)"
        }
        else {
            Write-Host "##[error]The request failed before receiving a response (network/DNS/TLS). Verify the Jira base URL is reachable from the agent."
        }
        return $null
    }

    Write-Host "Site URL rejected the API token (HTTP 401); checking whether this is a scoped token that requires the Atlassian platform gateway..." -ForegroundColor Yellow
    $cloudId = Get-JiraCloudId -BaseUrl $BaseUrl
    if (-not $cloudId) {
        Write-Host "##[error]Authentication failed: $BaseUrl rejected the token (HTTP 401), and cloud ID discovery via $BaseUrl/_edge/tenant_info failed, so the Atlassian platform gateway could not be tried."
        Write-Host "##[error]Verify the Jira email and API token, and that the base URL points at a Jira Cloud site."
        return $null
    }

    $gatewayRoot = "$GatewayBaseUrl/ex/jira/$cloudId"
    $gatewayAuth = Test-JiraAuth -Root $gatewayRoot -Headers $Headers
    if ($gatewayAuth.Ok) {
        Write-Host "Authenticated to Jira as: $($gatewayAuth.DisplayName)" -ForegroundColor Green
        Write-Host "Scoped API token detected — using the Atlassian platform gateway ($gatewayRoot)." -ForegroundColor Cyan
        return $gatewayRoot
    }

    $gatewayStatus = if ($gatewayAuth.StatusCode) { "HTTP $($gatewayAuth.StatusCode)" } else { 'no response' }
    Write-Host "##[error]Authentication failed at both the site URL and the Atlassian platform gateway."
    Write-Host "##[error]  Site URL: GET $BaseUrl/rest/api/3/myself -> HTTP 401"
    Write-Host "##[error]  Gateway:  GET $gatewayRoot/rest/api/3/myself -> $gatewayStatus"
    Write-Host "##[error]Atlassian issues two token types that look identical: classic API tokens work at the site URL; scoped API tokens (the only type service accounts can create) work only through the gateway and must include the classic scopes 'read:jira-work' (issue access) and 'read:jira-user' (identity check)."
    Write-Host "##[error]Verify the email/token pair, confirm a scoped token includes both scopes, and check that the token has not expired."
    return $null
}

function Resolve-JiraCustomFields {
    param(
        [string]$ApiRoot,
        [hashtable]$Headers,
        [string]$CustomFields
    )

    # Turns the configured custom-field list (display names and/or
    # customfield_NNNNN IDs) into a deduplicated list of field IDs. Display
    # names cost one GET /rest/api/3/field catalog call, made only when at
    # least one name needs resolving. Custom fields are supplemental review
    # context, so every failure here degrades with a warning instead of
    # aborting the run.
    if ([string]::IsNullOrWhiteSpace($CustomFields)) {
        return @()
    }

    $entries = @($CustomFields -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($entries.Count -eq 0) {
        return @()
    }

    $resolved = [System.Collections.Generic.List[string]]::new()
    $pendingNames = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $entries) {
        if ($entry -match '^(?i)customfield_\d+$') {
            $id = $entry.ToLowerInvariant()
            if (-not $resolved.Contains($id)) { $resolved.Add($id) }
        }
        else {
            $pendingNames.Add($entry)
        }
    }

    if ($pendingNames.Count -gt 0) {
        $catalog = $null
        try {
            $catalog = Invoke-RestMethod -Uri "$ApiRoot/rest/api/3/field" -Headers $Headers -Method Get -ErrorAction Stop
        }
        catch {
            Write-Host "##[warning]Could not retrieve the Jira field catalog to resolve custom field names ($($_.Exception.Message)); skipping: $($pendingNames -join ', ')"
        }

        if ($catalog) {
            # Case-insensitive display-name -> ID map. Jira allows several
            # fields to share a display name; the first one wins.
            $map = @{}
            $duplicates = @{}
            foreach ($field in $catalog) {
                $nameKey = ([string]$field.name).ToLowerInvariant()
                if (-not $nameKey) { continue }
                if ($map.ContainsKey($nameKey)) {
                    $duplicates[$nameKey] = $true
                }
                else {
                    $map[$nameKey] = $field.id
                }
            }

            foreach ($name in $pendingNames) {
                $nameKey = $name.ToLowerInvariant()
                if ($map.ContainsKey($nameKey)) {
                    if ($duplicates.ContainsKey($nameKey)) {
                        Write-Host "Note: multiple Jira fields are named '$name'; using the first ($($map[$nameKey]))." -ForegroundColor Yellow
                    }
                    if (-not $resolved.Contains($map[$nameKey])) { $resolved.Add($map[$nameKey]) }
                }
                else {
                    Write-Host "##[warning]Jira custom field '$name' not found on this site; skipping."
                }
            }
        }
    }

    return @($resolved)
}

function Format-JiraFieldValue {
    param($Value)

    # Renders a raw (non-HTML) field value as display text, covering the
    # common Jira field shapes: plain scalars, option objects ({value:...}),
    # named objects ({name:...}), user objects ({displayName:...}), and
    # arrays of any of those. Unrecognized objects (e.g. Atlassian Document
    # Format bodies) fall back to compact JSON rather than PowerShell's
    # lossy default ToString().
    if ($null -eq $Value) { return '' }

    if ($Value -is [array]) {
        $parts = foreach ($item in $Value) {
            $part = Format-JiraFieldValue -Value $item
            if (-not [string]::IsNullOrWhiteSpace($part)) { $part }
        }
        return (@($parts) -join ', ')
    }

    if ($Value -is [string]) { return $Value }
    if ($Value -is [ValueType]) { return "$Value" }

    if ($Value.PSObject.Properties['name'] -and $Value.name -is [string]) { return [string]$Value.name }
    if ($Value.PSObject.Properties['value'] -and ($Value.value -is [string] -or $Value.value -is [ValueType])) { return [string]$Value.value }
    if ($Value.PSObject.Properties['displayName'] -and $Value.displayName -is [string]) { return [string]$Value.displayName }

    try {
        return ($Value | ConvertTo-Json -Compress -Depth 5)
    }
    catch {
        return "$Value"
    }
}

function Get-JiraCommentWindow {
    param(
        [string]$ApiRoot,
        [hashtable]$Headers,
        [string]$Key,
        $Issue
    )

    # Returns up to the 10 most recent comments (oldest-first within that
    # window) as @{Author; Created; BodyHtml}, plus a header note. The
    # comment container embedded in the issue GET holds the OLDEST page when
    # the server caps it, so a capped container needs one follow-up request
    # windowed to the tail (ascending by created, startAt = total - 10);
    # renderedBody rides along per item there, so no cross-request index
    # alignment is required.
    $result = @{ Comments = @(); HeaderNote = '' }

    $container = $Issue.fields.comment
    if (-not $container -or -not $container.total -or [int]$container.total -eq 0) {
        return $result
    }

    $total = [int]$container.total
    $windowSize = [Math]::Min(10, $total)
    $result.HeaderNote = if ($total -gt 10) { "showing the 10 most recent of $total" } else { "$total" }

    $embedded = @($container.comments)
    if ($embedded.Count -ge $total) {
        # Complete set: rendered bodies align by index with the raw comments
        # (same request, same page).
        $rendered = @($Issue.renderedFields.comment.comments)
        $windowStart = $embedded.Count - $windowSize
        $window = for ($i = 0; $i -lt $windowSize; $i++) {
            $index = $windowStart + $i
            @{
                Author   = $embedded[$index].author.displayName
                Created  = $embedded[$index].created
                BodyHtml = if ($index -lt $rendered.Count) { $rendered[$index].body } else { $null }
            }
        }
        $result.Comments = @($window)
        return $result
    }

    try {
        $page = Invoke-RestMethod -Uri "$ApiRoot/rest/api/3/issue/$Key/comment?orderBy=created&startAt=$([Math]::Max(0, $total - 10))&maxResults=10&expand=renderedBody" -Headers $Headers -Method Get -ErrorAction Stop
        $window = foreach ($comment in @($page.comments)) {
            @{
                Author   = $comment.author.displayName
                Created  = $comment.created
                BodyHtml = $comment.renderedBody
            }
        }
        $result.Comments = @($window)
    }
    catch {
        # Comments are supplemental — degrade to the oldest available page
        # rather than failing the run.
        Write-Host "##[warning]Could not fetch the most recent comments for ${Key}: $($_.Exception.Message). Showing older comments instead."
        $rendered = @($Issue.renderedFields.comment.comments)
        $tailStart = [Math]::Max(0, $embedded.Count - $windowSize)
        $window = for ($i = $tailStart; $i -lt $embedded.Count; $i++) {
            @{
                Author   = $embedded[$i].author.displayName
                Created  = $embedded[$i].created
                BodyHtml = if ($i -lt $rendered.Count) { $rendered[$i].body } else { $null }
            }
        }
        $result.Comments = @($window)
        $result.HeaderNote = 'showing older comments; the most recent were unavailable'
    }

    return $result
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

# Refuse to send Basic auth credentials over cleartext: Jira Cloud is
# HTTPS-only, so any http:// target is a misconfiguration. Loopback addresses
# are exempt (traffic never leaves the host) so tests can run against a local
# mock API.
$baseUri = $null
if (-not [Uri]::TryCreate($BaseUrl, [UriKind]::Absolute, [ref]$baseUri)) {
    Write-Host "##[error]Invalid Jira base URL: $BaseUrl"
    exit 1
}
if ($baseUri.Scheme -ne 'https' -and -not $baseUri.IsLoopback) {
    Write-Host "##[error]Jira base URL must use HTTPS ($BaseUrl). Basic auth credentials must not be sent over cleartext HTTP."
    exit 1
}

# The gateway base URL is a test seam, but it carries the same credentials —
# hold it to the same https-or-loopback rule so it can never become a
# cleartext loophole.
$GatewayBaseUrl = $GatewayBaseUrl.TrimEnd('/')
$gatewayUri = $null
if (-not [Uri]::TryCreate($GatewayBaseUrl, [UriKind]::Absolute, [ref]$gatewayUri)) {
    Write-Host "##[error]Invalid gateway base URL: $GatewayBaseUrl"
    exit 1
}
if ($gatewayUri.Scheme -ne 'https' -and -not $gatewayUri.IsLoopback) {
    Write-Host "##[error]Gateway base URL must use HTTPS ($GatewayBaseUrl). Basic auth credentials must not be sent over cleartext HTTP."
    exit 1
}

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

# Resolve which API root this token authenticates against (site URL for
# classic tokens, platform gateway for scoped tokens) BEFORE fetching any
# issues. Issue GETs mask bad auth as 404 "not found", so without this
# preflight an auth failure is indistinguishable from a false-positive key.
# Runs after the zero-candidates exit above so no-op runs make no API calls.
$apiRoot = Resolve-JiraApiRoot -BaseUrl $BaseUrl -GatewayBaseUrl $GatewayBaseUrl -Headers $headers
if (-not $apiRoot) {
    exit 1
}

# Resolve configured custom fields (display names -> IDs) before building
# the issue query; failures degrade with warnings inside the resolver.
$customFieldIds = @(Resolve-JiraCustomFields -ApiRoot $apiRoot -Headers $headers -CustomFields $CustomFields)

# Fetch each candidate; renderedFields returns HTML for rich-text fields.
# 404s are skipped (false-positive keys); any other API failure aborts with
# exit 1 — diagnostics were already printed by Invoke-JiraApi.
$fields = "summary,issuetype,status,priority,description,labels,components,fixVersions,duedate,assignee,reporter,parent,subtasks,issuelinks,comment,environment"
if ($customFieldIds.Count -gt 0) {
    $fields += ',' + ($customFieldIds -join ',')
}
# The names expand maps field IDs to display names; only needed to label
# configured custom fields in the output.
$expand = if ($customFieldIds.Count -gt 0) { 'renderedFields,names' } else { 'renderedFields' }
$issues = @()
try {
    foreach ($key in $candidateKeys) {
        $issueUrl = "$apiRoot/rest/api/3/issue/$key`?fields=$fields&expand=$expand"
        $issue = Invoke-JiraApi -Uri $issueUrl -Headers $headers

        if ($null -eq $issue) {
            # Auth is proven good by the preflight, so a 404 here means the
            # key does not exist OR the account lacks Browse permission on
            # its project — Jira deliberately returns the same 404 for both.
            Write-Host "  $key — not found in Jira, or the account lacks Browse permission (skipping; possibly a false-positive match)" -ForegroundColor DarkGray
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

    # Single-line context fields — each emitted only when populated
    if ($issue.fields.assignee -and $issue.fields.assignee.displayName) {
        Write-Output-Line "  Assignee:        $($issue.fields.assignee.displayName)"
    }
    if ($issue.fields.reporter -and $issue.fields.reporter.displayName) {
        Write-Output-Line "  Reporter:        $($issue.fields.reporter.displayName)"
    }
    if ($issue.fields.duedate) {
        Write-Output-Line "  Due Date:        $($issue.fields.duedate)"
    }
    # Where-Object filters guard against absent fields: a missing property
    # reads as $null, and @($null) has Count 1, which would render an empty
    # section label.
    $labels = @($issue.fields.labels | Where-Object { $_ })
    if ($labels.Count -gt 0) {
        Write-Output-Line "  Labels:          $($labels -join ', ')"
    }
    $components = @($issue.fields.components | Where-Object { $_ } | ForEach-Object { $_.name })
    if ($components.Count -gt 0) {
        Write-Output-Line "  Components:      $($components -join ', ')"
    }
    $fixVersions = @($issue.fields.fixVersions | Where-Object { $_ } | ForEach-Object { $_.name })
    if ($fixVersions.Count -gt 0) {
        Write-Output-Line "  Fix Versions:    $($fixVersions -join ', ')"
    }
    if ($issue.fields.parent) {
        Write-Output-Line "  Parent:          $($issue.fields.parent.key) — $($issue.fields.parent.fields.summary)"
    }

    if (-not [string]::IsNullOrWhiteSpace($description)) {
        Write-Output-Line "`n  Description:"
        $description -split "`n" | ForEach-Object {
            Write-Output-Line "    $_"
        }
    }

    # Environment: classic rich-text field, mostly used for repro context
    $environment = ConvertFrom-Html $issue.renderedFields.environment
    if (-not [string]::IsNullOrWhiteSpace($environment)) {
        Write-Output-Line "`n  Environment:"
        $environment -split "`n" | ForEach-Object {
            Write-Output-Line "    $_"
        }
    }

    # Configured custom fields, in configured order. Rich-text values come
    # back rendered as HTML; everything else goes through the raw formatter.
    foreach ($fieldId in $customFieldIds) {
        $label = $null
        if ($issue.names) { $label = $issue.names.$fieldId }
        if (-not $label) { $label = $fieldId }

        $renderedValue = $null
        if ($issue.renderedFields) { $renderedValue = $issue.renderedFields.$fieldId }
        $value = if (-not [string]::IsNullOrWhiteSpace([string]$renderedValue)) {
            ConvertFrom-Html ([string]$renderedValue)
        }
        else {
            Format-JiraFieldValue -Value $issue.fields.$fieldId
        }
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        if ($value -match "`n") {
            Write-Output-Line "`n  ${label}:"
            $value -split "`n" | ForEach-Object {
                Write-Output-Line "    $_"
            }
        }
        else {
            # Leading blank line so a short custom field does not read as a
            # continuation of the preceding block (e.g. the description)
            Write-Output-Line "`n  ${label}: $value"
        }
    }

    $subtasks = @($issue.fields.subtasks | Where-Object { $_ })
    if ($subtasks.Count -gt 0) {
        Write-Output-Line "`n  Subtasks:"
        foreach ($subtask in $subtasks) {
            Write-Output-Line "    - $($subtask.key): $($subtask.fields.summary) [$($subtask.fields.status.name)]"
        }
    }

    # Issue links carry a directional verb on the type (e.g. "blocks" vs
    # "is blocked by"); emit whichever side of the link this issue is on.
    $linkLines = foreach ($link in @($issue.fields.issuelinks)) {
        if ($link.outwardIssue) {
            "- $($link.type.outward) $($link.outwardIssue.key): $($link.outwardIssue.fields.summary) [$($link.outwardIssue.fields.status.name)]"
        }
        elseif ($link.inwardIssue) {
            "- $($link.type.inward) $($link.inwardIssue.key): $($link.inwardIssue.fields.summary) [$($link.inwardIssue.fields.status.name)]"
        }
    }
    $linkLines = @($linkLines | Where-Object { $_ })
    if ($linkLines.Count -gt 0) {
        Write-Output-Line "`n  Linked Issues:"
        $linkLines | ForEach-Object { Write-Output-Line "    $_" }
    }

    # Discussion often carries the real acceptance nuance — cap at the 10
    # most recent so long-lived tickets cannot flood the review context.
    $commentWindow = Get-JiraCommentWindow -ApiRoot $apiRoot -Headers $headers -Key $issue.key -Issue $issue
    if (@($commentWindow.Comments).Count -gt 0) {
        Write-Output-Line "`n  Comments ($($commentWindow.HeaderNote)):"
        foreach ($comment in $commentWindow.Comments) {
            # Invoke-RestMethod deserializes ISO timestamps into [datetime]
            # (converted to local time); normalize back to the UTC date so
            # the display is stable regardless of the agent's timezone.
            $when = $comment.Created
            if ($when -is [datetime]) {
                $when = $when.ToUniversalTime().ToString('yyyy-MM-dd')
            }
            else {
                $when = [string]$when
                if ($when.Length -ge 10) { $when = $when.Substring(0, 10) }
            }
            Write-Output-Line "    [$($comment.Author) — $when]"
            $body = if ($comment.BodyHtml) { ConvertFrom-Html $comment.BodyHtml } else { '' }
            if ([string]::IsNullOrWhiteSpace($body)) { $body = '(comment body unavailable)' }
            $body -split "`n" | ForEach-Object {
                Write-Output-Line "      $_"
            }
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
