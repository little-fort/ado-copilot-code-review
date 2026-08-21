<#
    Tests for the Jira work-item integration (issue #53).

    Key extraction is unit-tested by extracting Get-JiraIssueKeys from the
    script via the PowerShell AST. The fetch flow runs the real script as a
    child pwsh process against a local HttpListener standing in for the Jira
    Cloud REST API.

    Requires Pester 5+. Run via: pwsh -File tests/powershell/Invoke-Tests.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:JiraScript = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Get-JiraWorkItems.ps1'

    # Shared mock-API harness (HttpListener + child pwsh process)
    . (Join-Path $PSScriptRoot "MockApiHelper.ps1")

    # Extract Get-JiraIssueKeys for direct unit testing without executing the script
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($JiraScript, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw "Get-JiraWorkItems.ps1 failed to parse: $($parseErrors -join '; ')"
    }
    $funcAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Get-JiraIssueKeys'
    }, $true)
    if (-not $funcAst) {
        throw 'Get-JiraIssueKeys function not found in Get-JiraWorkItems.ps1 — update this test if it was renamed.'
    }
    . ([scriptblock]::Create($funcAst.Extent.Text))

    # Same extraction for the raw field-value formatter (commit 2)
    $fmtAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Format-JiraFieldValue'
    }, $true)
    if (-not $fmtAst) {
        throw 'Format-JiraFieldValue function not found in Get-JiraWorkItems.ps1 — update this test if it was renamed.'
    }
    . ([scriptblock]::Create($fmtAst.Extent.Text))

    # Writes a PR_Metadata.json into a fresh temp dir and returns its path
    function New-PrMetadataFile {
        param([hashtable]$Metadata)
        $dir = Join-Path ([System.IO.Path]::GetTempPath()) ("jira-test-" + [guid]::NewGuid())
        New-Item -ItemType Directory -Path $dir | Out-Null
        $file = Join-Path $dir 'PR_Metadata.json'
        $Metadata | ConvertTo-Json | Out-File -FilePath $file -Encoding UTF8
        return $file
    }

    $script:IssueJson = @'
{"key":"PROJ-1","fields":{"summary":"Add login flow","issuetype":{"name":"Story"},"status":{"name":"In Progress"},"priority":{"name":"High"}},"renderedFields":{"description":"<p>Login must support SSO.</p><p><del>Old requirement</del></p>"}}
'@

    # Auth preflight (GET /rest/api/3/myself) and cloud ID discovery fixtures
    $script:MyselfJson = @'
{"displayName":"CI Bot","emailAddress":"me@example.com"}
'@
    $script:TenantInfoJson = @'
{"cloudId":"tid-123"}
'@

    # Rich issue exercising every expanded field (commit 2). fixVersions is
    # deliberately empty to prove empty sections are omitted; the custom
    # field's raw value is null so its rendered HTML must be used instead.
    $script:RichIssueJson = @'
{"key":"PROJ-1","names":{"customfield_10031":"Acceptance Criteria"},"fields":{"summary":"Add login flow","issuetype":{"name":"Story"},"status":{"name":"In Progress"},"priority":{"name":"High"},"labels":["auth","security"],"components":[{"name":"WebApp"}],"fixVersions":[],"duedate":"2026-09-01","assignee":{"displayName":"Dana Dev"},"reporter":{"displayName":"Rae Reporter"},"parent":{"key":"PROJ-100","fields":{"summary":"Auth epic"}},"subtasks":[{"key":"PROJ-2","fields":{"summary":"Add SSO config","status":{"name":"To Do"}}}],"issuelinks":[{"type":{"outward":"blocks","inward":"is blocked by"},"outwardIssue":{"key":"PROJ-9","fields":{"summary":"Deploy login","status":{"name":"Open"}}}},{"type":{"outward":"relates to","inward":"relates to"},"inwardIssue":{"key":"CORE-3","fields":{"summary":"Session store","status":{"name":"Done"}}}}],"comment":{"total":3,"comments":[{"author":{"displayName":"Alice"},"created":"2026-08-01T10:00:00.000+0000"},{"author":{"displayName":"Bob"},"created":"2026-08-02T10:00:00.000+0000"},{"author":{"displayName":"Carol"},"created":"2026-08-03T10:00:00.000+0000"}]},"customfield_10031":null},"renderedFields":{"description":"<p>Login must support SSO.</p>","environment":"<p>Staging cluster only</p>","customfield_10031":"<ul><li>Given a user, SSO login succeeds</li></ul>","comment":{"comments":[{"body":"<p>First comment</p>"},{"body":"<p>Second comment</p>"},{"body":"<p>Third <b>rich</b> comment</p>"}]}}}
'@

    $script:FieldCatalogJson = @'
[{"id":"summary","name":"Summary","custom":false},{"id":"customfield_10031","name":"Acceptance Criteria","custom":true},{"id":"customfield_10032","name":"Decoy Field","custom":true}]
'@

    # Comment-heavy fixtures are built programmatically — hand-writing 20+
    # comment objects as literals would be unreadable.
    $newComment = { param($n) @{ author = @{ displayName = "C$n" }; created = ('2026-08-{0:d2}T00:00:00.000+0000' -f $n) } }
    $newRendered = { param($n) @{ body = "<p>Comment number $n</p>" } }

    # 12 comments, all embedded (total == returned): no follow-up request
    $script:TwelveCommentIssueJson = [ordered]@{
        key            = 'PROJ-1'
        fields         = [ordered]@{
            summary   = 'Busy ticket'
            issuetype = @{ name = 'Task' }
            status    = @{ name = 'Open' }
            comment   = @{ total = 12; comments = @(1..12 | ForEach-Object { & $newComment $_ }) }
        }
        renderedFields = @{ comment = @{ comments = @(1..12 | ForEach-Object { & $newRendered $_ }) } }
    } | ConvertTo-Json -Depth 10 -Compress

    # 25 comments but only the oldest 20 embedded (server-capped): the tail
    # must come from a windowed follow-up request
    $script:CappedCommentIssueJson = [ordered]@{
        key            = 'PROJ-1'
        fields         = [ordered]@{
            summary   = 'Very busy ticket'
            issuetype = @{ name = 'Task' }
            status    = @{ name = 'Open' }
            comment   = @{ total = 25; comments = @(1..20 | ForEach-Object { & $newComment $_ }) }
        }
        renderedFields = @{ comment = @{ comments = @(1..20 | ForEach-Object { & $newRendered $_ }) } }
    } | ConvertTo-Json -Depth 10 -Compress

    $script:CommentPageJson = @{
        comments = @(16..25 | ForEach-Object {
            @{ author = @{ displayName = "C$_" }; created = ('2026-08-{0:d2}T00:00:00.000+0000' -f $_); renderedBody = "<p>Comment number $_</p>" }
        })
    } | ConvertTo-Json -Depth 10 -Compress
}

Describe 'Get-JiraIssueKeys' {

    It 'extracts keys from branch names, titles, and descriptions' {
        $keys = Get-JiraIssueKeys -Text "refs/heads/feature/PROJ-123-login`nFix CORE-9 regression`nRelates to PROJ-456"
        $keys | Should -Be @('PROJ-123', 'CORE-9', 'PROJ-456')
    }

    It 'uppercases lowercase keys from branch names' {
        Get-JiraIssueKeys -Text 'feature/proj-42-fix-thing' | Should -Be @('PROJ-42')
    }

    It 'deduplicates while preserving discovery order' {
        Get-JiraIssueKeys -Text 'PROJ-1 then CORE-2 then PROJ-1 again' | Should -Be @('PROJ-1', 'CORE-2')
    }

    It 'filters by project keys when provided' {
        $keys = Get-JiraIssueKeys -Text 'PROJ-1 CORE-2 OTHER-3' -ProjectKeys 'proj, core'
        $keys | Should -Be @('PROJ-1', 'CORE-2')
    }

    It 'caps the number of returned keys' {
        $text = (1..20 | ForEach-Object { "PROJ-$_" }) -join ' '
        (Get-JiraIssueKeys -Text $text -MaxKeys 10).Count | Should -Be 10
    }

    It 'returns nothing for empty input' {
        Get-JiraIssueKeys -Text '' | Should -BeNullOrEmpty
        Get-JiraIssueKeys -Text 'no keys here, just words-and-hyphens' | Should -BeNullOrEmpty
    }

    It 'does not match single-letter prefixes or trailing-letter tokens' {
        # 'a-1' (single letter) and 'abc-123def' (digits followed by letters) are not key-shaped
        Get-JiraIssueKeys -Text 'a-1 abc-123def' | Should -BeNullOrEmpty
    }
}

Describe 'Get-JiraWorkItems.ps1 fetch flow (issue #53)' {

    It 'fetches real issues, skips 404 false positives, and converts HTML descriptions' {
        # Branch yields PROJ-1 (real); title yields UTF-8 (false positive → 404)
        $metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/PROJ-1-login'
            title         = 'Add login (UTF-8 support)'
            description   = $null
        }
        $outputFile = Join-Path (Split-Path -Parent $metadataFile) 'Work_Item_Details.txt'

        try {
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 200; Body = $MyselfJson },
                @{ Status = 200; Body = $IssueJson },
                @{ Status = 404; Body = '{"errorMessages":["Issue does not exist"]}' }
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
                  '-PrMetadataFile', $metadataFile, '-OutputFile', $outputFile)
            }

            $result.ExitCode | Should -Be 0
            # Auth preflight first, then both candidates in discovery order
            $result.Requests.Count | Should -Be 3
            $result.Requests[0].Path | Should -Match '/rest/api/3/myself'
            $result.Requests[1].Path | Should -Match 'PROJ-1'
            $result.Requests[2].Path | Should -Match 'UTF-8'
            # Classic-token regression: a passing site preflight must never
            # touch cloud ID discovery or the platform gateway
            $result.Requests | Where-Object { $_.Path -match '_edge|/ex/jira/' } | Should -BeNullOrEmpty
            $result.Output | Should -Match 'Authenticated to Jira as: CI Bot'
            # The false positive was skipped without failing the run
            $result.Output | Should -Match 'UTF-8 — not found in Jira'
            # The real issue rendered with converted HTML (incl. strikethrough preservation)
            $result.Output | Should -Match '\[PROJ-1 - Story\]'
            $result.Output | Should -Match 'Login must support SSO\.'
            $result.Output | Should -Match '~~Old requirement~~'

            Test-Path $outputFile | Should -BeTrue
            Get-Content $outputFile -Raw | Should -Match '\[PROJ-1 - Story\]'
        }
        finally {
            Remove-Item (Split-Path -Parent $metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'exits cleanly without output when the PR references no issue keys' {
        $metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/add-login'
            title         = 'Add login'
            description   = 'No ticket references here'
        }

        try {
            # No responses needed — the script must not call the API at all
            $result = Invoke-ScriptWithMockApi -Responses @() -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
                  '-PrMetadataFile', $metadataFile)
            }

            $result.ExitCode | Should -Be 0
            $result.Requests.Count | Should -Be 0
            $result.Output | Should -Match 'No Jira issue keys found'
        }
        finally {
            Remove-Item (Split-Path -Parent $metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'exits non-zero with token-type and scope diagnostics when both auth preflights fail' {
        $metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/PROJ-1-login'
            title         = 'Add login'
            description   = $null
        }

        try {
            # Site /myself 401 → tenant_info discovery → gateway /myself 401
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 401; Body = '{"errorMessages":["Unauthorized"]}' },
                @{ Status = 200; Body = $TenantInfoJson },
                @{ Status = 401; Body = '{"code":401,"message":"Unauthorized"}' }
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'bad-token',
                  '-PrMetadataFile', $metadataFile, '-GatewayBaseUrl', $baseUrl)
            }

            $result.ExitCode | Should -Be 1
            $result.Requests.Count | Should -Be 3
            $result.Requests[0].Path | Should -Match '/rest/api/3/myself'
            $result.Requests[1].Path | Should -Match '_edge/tenant_info'
            $result.Requests[2].Path | Should -Match '/ex/jira/tid-123/rest/api/3/myself'
            $result.Output | Should -Match 'Authentication failed'
            $result.Output | Should -Match '-> HTTP 401'
            # The diagnostics must name both required scopes for scoped tokens
            $result.Output | Should -Match 'read:jira-work'
            $result.Output | Should -Match 'read:jira-user'
        }
        finally {
            Remove-Item (Split-Path -Parent $metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'falls back to the platform gateway when the site URL rejects a scoped token' {
        $metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/PROJ-1-login'
            title         = 'Add login'
            description   = $null
        }

        try {
            # Scoped-token signature: site 401, then discovery + gateway succeed
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 401; Body = '{"errorMessages":["Unauthorized"]}' },
                @{ Status = 200; Body = $TenantInfoJson },
                @{ Status = 200; Body = $MyselfJson },
                @{ Status = 200; Body = $IssueJson }
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'scoped-token',
                  '-PrMetadataFile', $metadataFile, '-GatewayBaseUrl', $baseUrl)
            }

            $result.ExitCode | Should -Be 0
            $result.Requests.Count | Should -Be 4
            $result.Requests[1].Path | Should -Match '_edge/tenant_info'
            $result.Requests[2].Path | Should -Match '/ex/jira/tid-123/rest/api/3/myself'
            # Issue fetches must route through the gateway root, not the site URL
            $result.Requests[3].Path | Should -Match '/ex/jira/tid-123/rest/api/3/issue/PROJ-1'
            $result.Output | Should -Match 'Scoped API token detected'
            $result.Output | Should -Match '\[PROJ-1 - Story\]'
        }
        finally {
            Remove-Item (Split-Path -Parent $metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'fails with diagnostics when the site rejects the token and cloud ID discovery fails' {
        $metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/PROJ-1-login'
            title         = 'Add login'
            description   = $null
        }

        try {
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 401; Body = '{"errorMessages":["Unauthorized"]}' },
                @{ Status = 500; Body = 'oops' }
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'bad-token',
                  '-PrMetadataFile', $metadataFile, '-GatewayBaseUrl', $baseUrl)
            }

            $result.ExitCode | Should -Be 1
            $result.Requests.Count | Should -Be 2
            $result.Output | Should -Match 'cloud ID discovery'
            $result.Output | Should -Match 'Authentication failed'
        }
        finally {
            Remove-Item (Split-Path -Parent $metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'refuses a non-loopback http gateway URL before sending credentials' {
        $result = Invoke-ScriptWithMockApi -Responses @() -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', 'unused.json', '-GatewayBaseUrl', 'http://gateway.internal.example.com')
        }

        $result.ExitCode | Should -Be 1
        $result.Output | Should -Match 'must use HTTPS'
        $result.Requests.Count | Should -Be 0
    }

    It 'refuses a non-loopback http base URL before sending credentials' {
        # Basic auth over cleartext must be rejected up front; the loopback
        # exemption exists only so these tests can use a local mock API
        $result = Invoke-ScriptWithMockApi -Responses @() -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', 'http://jira.internal.example.com', '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', 'unused.json')
        }

        $result.ExitCode | Should -Be 1
        $result.Output | Should -Match 'must use HTTPS'
        $result.Requests.Count | Should -Be 0
    }

    It 'exits non-zero when the metadata file is missing' {
        $result = Invoke-ScriptWithMockApi -Responses @() -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', 'Z:\does\not\exist.json')
        }

        $result.ExitCode | Should -Be 1
        $result.Output | Should -Match 'PR metadata file not found'
    }
}

Describe 'Format-JiraFieldValue' {

    It 'passes strings and numbers through' {
        Format-JiraFieldValue -Value 'plain text' | Should -Be 'plain text'
        Format-JiraFieldValue -Value 42 | Should -Be '42'
    }

    It 'returns empty for null' {
        Format-JiraFieldValue -Value $null | Should -Be ''
    }

    It 'joins arrays of named objects' {
        $value = @([pscustomobject]@{ name = 'WebApp' }, [pscustomobject]@{ name = 'API' })
        Format-JiraFieldValue -Value $value | Should -Be 'WebApp, API'
    }

    It 'joins plain string arrays' {
        Format-JiraFieldValue -Value @('auth', 'security') | Should -Be 'auth, security'
    }

    It 'unwraps option objects and user objects' {
        Format-JiraFieldValue -Value ([pscustomobject]@{ value = 'High' }) | Should -Be 'High'
        Format-JiraFieldValue -Value ([pscustomobject]@{ displayName = 'Dana Dev' }) | Should -Be 'Dana Dev'
    }

    It 'falls back to compact JSON for unrecognized objects (e.g. ADF bodies)' {
        $adf = [pscustomobject]@{ type = 'doc'; version = 1 }
        Format-JiraFieldValue -Value $adf | Should -Match '"type":\s*"doc"'
    }
}

Describe 'Get-JiraWorkItems.ps1 expanded fields' {

    BeforeEach {
        $script:metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/PROJ-1-login'
            title         = 'Add login'
            description   = $null
        }
    }

    AfterEach {
        Remove-Item (Split-Path -Parent $script:metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'renders the expanded field set and omits empty fields' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $MyselfJson },
            @{ Status = 200; Body = $RichIssueJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', $metadataFile)
        }

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 2
        # Pinned header format must survive the expansion
        $result.Output | Should -Match '\[PROJ-1 - Story\]'
        $result.Output | Should -Match 'Assignee:\s+Dana Dev'
        $result.Output | Should -Match 'Reporter:\s+Rae Reporter'
        $result.Output | Should -Match 'Due Date:\s+2026-09-01'
        $result.Output | Should -Match 'Labels:\s+auth, security'
        $result.Output | Should -Match 'Components:\s+WebApp'
        $result.Output | Should -Match 'Parent:\s+PROJ-100 — Auth epic'
        $result.Output | Should -Match 'Environment:'
        $result.Output | Should -Match 'Staging cluster only'
        $result.Output | Should -Match '- PROJ-2: Add SSO config \[To Do\]'
        $result.Output | Should -Match '- blocks PROJ-9: Deploy login \[Open\]'
        $result.Output | Should -Match '- relates to CORE-3: Session store \[Done\]'
        $result.Output | Should -Match 'Comments \(3\):'
        $result.Output | Should -Match '\[Alice — 2026-08-01\]'
        $result.Output | Should -Match 'Third rich comment'
        # Empty fields must not produce section labels
        $result.Output | Should -Not -Match 'Fix Versions:'
        # Custom fields render only when configured
        $result.Output | Should -Not -Match 'Acceptance Criteria'
    }

    It 'caps comments at the 10 most recent when the embedded set is complete' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $MyselfJson },
            @{ Status = 200; Body = $TwelveCommentIssueJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', $metadataFile)
        }

        $result.ExitCode | Should -Be 0
        # Complete embedded set: no follow-up comment request
        $result.Requests.Count | Should -Be 2
        $result.Output | Should -Match 'Comments \(showing the 10 most recent of 12\):'
        $result.Output | Should -Match '\[C3 — 2026-08-03\]'
        $result.Output | Should -Match '\[C12 — 2026-08-12\]'
        $result.Output | Should -Not -Match '\[C2 —'
    }

    It 'fetches the comment tail when the embedded page is server-capped' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $MyselfJson },
            @{ Status = 200; Body = $CappedCommentIssueJson },
            @{ Status = 200; Body = $CommentPageJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', $metadataFile)
        }

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 3
        # The follow-up must window to the tail, ascending, with rendered bodies
        $result.Requests[2].Path | Should -Match '/rest/api/3/issue/PROJ-1/comment\?'
        $result.Requests[2].Path | Should -Match 'orderBy=created'
        $result.Requests[2].Path | Should -Match 'startAt=15'
        $result.Requests[2].Path | Should -Match 'maxResults=10'
        $result.Requests[2].Path | Should -Match 'expand=renderedBody'
        $result.Output | Should -Match 'Comments \(showing the 10 most recent of 25\):'
        $result.Output | Should -Match '\[C25 — 2026-08-25\]'
        $result.Output | Should -Not -Match '\[C5 —'
    }

    It 'includes a custom field by ID without a catalog call' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $MyselfJson },
            @{ Status = 200; Body = $RichIssueJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', $metadataFile, '-CustomFields', 'customfield_10031')
        }

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 2
        $result.Requests | Where-Object { $_.Path -match '/rest/api/3/field' } | Should -BeNullOrEmpty
        # The issue query must request the field and the names expand
        $result.Requests[1].Path | Should -Match 'customfield_10031'
        $result.Requests[1].Path | Should -Match 'names'
        # Section label comes from the names map; value from renderedFields
        $result.Output | Should -Match 'Acceptance Criteria: Given a user, SSO login succeeds'
    }

    It 'resolves a custom field display name through the field catalog' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $MyselfJson },
            @{ Status = 200; Body = $FieldCatalogJson },
            @{ Status = 200; Body = $RichIssueJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', $metadataFile, '-CustomFields', 'acceptance criteria')
        }

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 3
        $result.Requests[1].Path | Should -Match '/rest/api/3/field'
        $result.Requests[2].Path | Should -Match 'customfield_10031'
        $result.Output | Should -Match 'Acceptance Criteria: Given a user, SSO login succeeds'
    }

    It 'warns and continues when a custom field name cannot be resolved' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $MyselfJson },
            @{ Status = 200; Body = $FieldCatalogJson },
            @{ Status = 200; Body = $IssueJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $JiraScript,
              '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
              '-PrMetadataFile', $metadataFile, '-CustomFields', 'Nope Field')
        }

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match "warning\]Jira custom field 'Nope Field' not found"
        # The review context still renders without the unresolvable field
        $result.Output | Should -Match '\[PROJ-1 - Story\]'
    }
}
