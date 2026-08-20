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
                @{ Status = 200; Body = $IssueJson },
                @{ Status = 404; Body = '{"errorMessages":["Issue does not exist"]}' }
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'test-token',
                  '-PrMetadataFile', $metadataFile, '-OutputFile', $outputFile)
            }

            $result.ExitCode | Should -Be 0
            # Both candidates were tried, in discovery order
            $result.Requests.Count | Should -Be 2
            $result.Requests[0].Path | Should -Match 'PROJ-1'
            $result.Requests[1].Path | Should -Match 'UTF-8'
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

    It 'exits non-zero with diagnostics on authentication failure' {
        $metadataFile = New-PrMetadataFile -Metadata @{
            sourceRefName = 'refs/heads/feature/PROJ-1-login'
            title         = 'Add login'
            description   = $null
        }

        try {
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 401; Body = '{"errorMessages":["Unauthorized"]}' }
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $JiraScript,
                  '-BaseUrl', $baseUrl, '-Email', 'me@example.com', '-ApiToken', 'bad-token',
                  '-PrMetadataFile', $metadataFile)
            }

            $result.ExitCode | Should -Be 1
            $result.Output | Should -Match 'HTTP status: 401'
            $result.Output | Should -Match 'Authentication failed'
        }
        finally {
            Remove-Item (Split-Path -Parent $metadataFile) -Recurse -Force -ErrorAction SilentlyContinue
        }
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
