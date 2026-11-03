<#
    Tests for the script side of incremental re-reviews (issue #67):

    - Set-AzureDevOpsPRReviewState.ps1 records a reviewed iteration as a single
      idempotent JSON Patch on the pull request properties.
    - The comment scripts leave "<id>.started"/"<id>.done" activity markers so
      the task can tell whether every posting attempt went through.
    - The context scripts leave "*.degraded" markers when part of the review
      context failed to load although the script itself succeeded.
    - The bookkeeping helpers stay identical across the scripts that define them.

    Scripts run as real child pwsh processes against a local HttpListener
    (MockApiHelper.ps1). Requires Pester 5+. Run via tests/powershell/Invoke-Tests.ps1.
#>

BeforeDiscovery {
    $script:ActivityHelperScripts = @('Add-CopilotComment.ps1', 'Add-AzureDevOpsPRComment.ps1', 'Update-CopilotComment.ps1', 'Delete-CopilotComment.ps1')
    $script:ContextHelperScripts = @('Get-AzureDevOpsPR.ps1', 'Get-AzureDevOpsWorkItems.ps1', 'Get-JiraWorkItems.ps1')
}

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:ScriptsDir = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts'
    . (Join-Path $PSScriptRoot 'MockApiHelper.ps1')

    # Source text of a function in a script (line endings normalized), or $null
    function Get-FunctionText {
        param([string]$ScriptName, [string]$FunctionName)
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $ScriptsDir $ScriptName), [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            throw "$ScriptName failed to parse: $($parseErrors -join '; ')"
        }
        $funcAst = $ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
        }, $true)
        if ($funcAst) { $funcAst.Extent.Text -replace "`r`n", "`n" } else { $null }
    }

    # A fresh per-test state folder, like the one the task creates per run
    function New-StateDir {
        $dir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $dir | Out-Null
        $dir
    }

    # The review policy the task writes into the state folder before the agent
    # runs (PR #65 rework): without it the comment scripts refuse every action.
    # Thread 42 is an automated-review thread the agent may act on.
    function Get-PolicyFiles {
        param([string]$StateDir, [string]$WorkspaceRoot = $StateDir)
        @{ (Join-Path $StateDir 'review-policy.json') = (New-TestPolicyJson -WorkspaceRoot $WorkspaceRoot -Threads @(New-TestPolicyThread -ThreadId 42 -AgentCommentIds 1)) }
    }

    # Activity markers as a summary: attempts started, attempts finished, and
    # the actions recorded by the finished ones
    function Get-ActivitySummary {
        param([string]$StateDir)
        $activityDir = Join-Path $StateDir 'activity'
        $files = if (Test-Path $activityDir) { @(Get-ChildItem -Path $activityDir -File) } else { @() }
        $started = @($files | Where-Object { $_.Extension -eq '.started' })
        $done = @($files | Where-Object { $_.Extension -eq '.done' })
        [pscustomobject]@{
            Started    = $started.Count
            Done       = $done.Count
            Unfinished = @($started | Where-Object { -not (Test-Path ([IO.Path]::ChangeExtension($_.FullName, '.done'))) }).Count
            Actions    = @($done | ForEach-Object { (Get-Content -Path $_.FullName -Raw).Trim() } | Sort-Object)
        }
    }

    $script:Sha = 'abcdef0123456789abcdef0123456789abcdef01'
    $script:PrefixValue = 'CopilotCodeReview.v1.0123456789abcdef'
    $script:PrJson = '{"title":"Test PR"}'
    $script:ThreadJson = '{"id":42,"status":"active","comments":[{"id":1,"author":{"displayName":"CI Mock"},"publishedDate":"2026-01-01T00:00:00Z"}]}'
    $script:CommentEnvironment = @{
        AZUREDEVOPS_TOKEN          = 'test-token'
        AZUREDEVOPS_AUTH_TYPE      = 'Basic'
        AZUREDEVOPS_COLLECTION_URI = '__BASEURL__/testorg'
        PROJECT                    = 'TestProject'
        REPOSITORY                 = 'TestRepo'
        PRID                       = '123'
    }
}

Describe 'Bookkeeping helper parity' {

    It 'defines Write-ReviewActivity identically in <_>' -ForEach $ActivityHelperScripts {
        $text = Get-FunctionText $_ 'Write-ReviewActivity'
        $text | Should -Not -BeNullOrEmpty
        $text | Should -BeExactly (Get-FunctionText 'Add-AzureDevOpsPRComment.ps1' 'Write-ReviewActivity')
    }

    It 'defines Write-ContextDegradation identically in <_>' -ForEach $ContextHelperScripts {
        $text = Get-FunctionText $_ 'Write-ContextDegradation'
        $text | Should -Not -BeNullOrEmpty
        $text | Should -BeExactly (Get-FunctionText 'Get-AzureDevOpsPR.ps1' 'Write-ContextDegradation')
    }
}

Describe 'Set-AzureDevOpsPRReviewState.ps1' {

    BeforeAll {
        $script:SetScript = Join-Path $ScriptsDir 'Set-AzureDevOpsPRReviewState.ps1'

        function Invoke-SetState {
            param([hashtable[]]$Responses, [string]$Iteration = '4', [string]$Commit = $Sha, [string]$Prefix = $PrefixValue)
            $setScript = $SetScript
            Invoke-ScriptWithMockApi -Responses $Responses -BuildArgs {
                param($baseUrl)
                @('-File', $setScript,
                  '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
                  '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '123',
                  '-StateKeyPrefix', $Prefix, '-Iteration', $Iteration, '-SourceCommit', $Commit)
            }.GetNewClosure()
        }
    }

    It 'records the iteration with exactly one JSON Patch request' {
        $result = Invoke-SetState -Responses @(@{ Status = 200; Body = '{"count":1,"value":{}}' })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 1
        $request = $result.Requests[0]
        $request.Method | Should -Be 'PATCH'
        $request.Path | Should -Match '/pullRequests/123/properties\?api-version=7\.1$'
        $request.ContentType | Should -Match '^application/json-patch\+json'

        # A one-operation patch must still be a JSON array
        $request.Body.TrimStart() | Should -Match '^\['
        $operations = @(ConvertFrom-Json -InputObject $request.Body -NoEnumerate)[0]
        @($operations).Count | Should -Be 1
        $operations[0].op | Should -Be 'replace'
        $operations[0].path | Should -Be "/$PrefixValue.Iteration.4"
        $operations[0].value | Should -Be $Sha
    }

    It 'retries a transient failure' {
        $result = Invoke-SetState -Responses @(
            @{ Status = 503; Body = '{"message":"unavailable"}' },
            @{ Status = 200; Body = '{"count":1,"value":{}}' }
        )
        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 2
    }

    It 'exits non-zero with diagnostics when the update is not permitted' {
        $result = Invoke-SetState -Responses @(@{ Status = 403; Body = '{"message":"TF401027: You need the Git Contribute permission"}' })
        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 1
        $result.Output | Should -Match '##\[error\]HTTP status: 403'
        $result.Output | Should -Match 'Contribute to pull requests'
    }

    It 'rejects an invalid <Name> before making any request' -ForEach @(
        @{ Name = 'commit SHA'; Iteration = '4'; Commit = 'not-a-sha'; Prefix = 'CopilotCodeReview.v1.0123456789abcdef' },
        @{ Name = 'iteration'; Iteration = '0'; Commit = 'abcdef0123456789abcdef0123456789abcdef01'; Prefix = 'CopilotCodeReview.v1.0123456789abcdef' },
        @{ Name = 'key prefix'; Iteration = '4'; Commit = 'abcdef0123456789abcdef0123456789abcdef01'; Prefix = 'bad/prefix' }
    ) {
        $result = Invoke-SetState -Responses @() -Iteration $Iteration -Commit $Commit -Prefix $Prefix
        $result.ExitCode | Should -Not -Be 0
        $result.Requests.Count | Should -Be 0
    }
}

Describe 'Posting activity evidence' {

    It 'records a finished attempt for a posted comment thread' {
        $stateDir = New-StateDir
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $ThreadJson }
        ) -Files (Get-PolicyFiles $stateDir) -Environment @{ COPILOT_REVIEW_STATE_DIR = $stateDir } -BuildArgs {
            param($baseUrl)
            @('-File', (Join-Path $ScriptsDir 'Add-AzureDevOpsPRComment.ps1'),
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '123', '-Comment', 'Feedback')
        }

        $result.ExitCode | Should -Be 0
        $summary = Get-ActivitySummary $stateDir
        $summary.Unfinished | Should -Be 0
        $summary.Actions | Should -Be @('thread')
    }

    It 'leaves an unfinished attempt when posting fails' {
        $stateDir = New-StateDir
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 400; Body = '{"message":"bad request"}' }
        ) -Files (Get-PolicyFiles $stateDir) -Environment @{ COPILOT_REVIEW_STATE_DIR = $stateDir } -BuildArgs {
            param($baseUrl)
            @('-File', (Join-Path $ScriptsDir 'Add-AzureDevOpsPRComment.ps1'),
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '123', '-Comment', 'Feedback')
        }

        $result.ExitCode | Should -Be 1
        # The API rejected the post, so this is not a policy refusal
        $result.Requests.Count | Should -Be 2
        $summary = Get-ActivitySummary $stateDir
        $summary.Started | Should -Be 1
        $summary.Unfinished | Should -Be 1
    }

    It 'records the wrapper and the inner script for a successful reply' {
        $stateDir = New-StateDir
        $workDir = New-StateDir
        Set-Content -LiteralPath (Join-Path $workDir '_reply.md') -Value 'Addressed.'
        $environment = $CommentEnvironment.Clone()
        $environment.COPILOT_REVIEW_STATE_DIR = $stateDir
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $ThreadJson },
            @{ Status = 200; Body = '{"id":2,"author":{"displayName":"CI Mock"},"publishedDate":"2026-01-01T00:00:00Z"}' }
        ) -Files (Get-PolicyFiles $stateDir $workDir) -WorkingDirectory $workDir -Environment $environment -BuildArgs {
            param($baseUrl)
            @('-File', (Join-Path $ScriptsDir 'Add-CopilotComment.ps1'), '-CommentFile', './_reply.md', '-ThreadId', '42')
        }

        $result.ExitCode | Should -Be 0
        $summary = Get-ActivitySummary $stateDir
        $summary.Unfinished | Should -Be 0
        $summary.Actions | Should -Be @('reply', 'wrapper')
    }

    It 'leaves an unfinished attempt when the wrapper exits early: <Case>' -ForEach @(
        @{ Case = 'missing comment file'; Arguments = @('-CommentFile', './does-not-exist.md') },
        @{ Case = 'invalid status value'; Arguments = @('-Comment', 'Feedback', '-Status', 'Fixed') },
        @{ Case = 'non-numeric line number'; Arguments = @('-Comment', 'Feedback', '-FilePath', '/src/a.cs', '-StartLine', '42-45') },
        @{ Case = 'zero thread ID'; Arguments = @('-Comment', 'Feedback', '-ThreadId', '0') }
    ) {
        $stateDir = New-StateDir
        $workDir = New-StateDir
        $environment = $CommentEnvironment.Clone()
        $environment.COPILOT_REVIEW_STATE_DIR = $stateDir
        # GetNewClosure captures only locals, so resolve the path here
        $wrapperArgs = $Arguments
        $wrapperScript = Join-Path $ScriptsDir 'Add-CopilotComment.ps1'
        $result = Invoke-ScriptWithMockApi -Responses @() -Files (Get-PolicyFiles $stateDir $workDir) -WorkingDirectory $workDir -Environment $environment -BuildArgs {
            param($baseUrl)
            @('-File', $wrapperScript) + $wrapperArgs
        }.GetNewClosure()

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $summary = Get-ActivitySummary $stateDir
        $summary.Started | Should -Be 1
        $summary.Unfinished | Should -Be 1
    }

    It 'records a finished update when a thread is resolved' {
        $stateDir = New-StateDir
        $environment = $CommentEnvironment.Clone()
        $environment.COPILOT_REVIEW_STATE_DIR = $stateDir
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = '{"id":42,"status":"fixed"}' }
        ) -Files (Get-PolicyFiles $stateDir) -Environment $environment -BuildArgs {
            param($baseUrl)
            @('-File', (Join-Path $ScriptsDir 'Update-CopilotComment.ps1'), '-ThreadId', '42', '-Status', 'Fixed')
        }

        $result.ExitCode | Should -Be 0
        $result.Requests[0].Method | Should -Be 'PATCH'
        $summary = Get-ActivitySummary $stateDir
        $summary.Unfinished | Should -Be 0
        $summary.Actions | Should -Be @('update')
    }

    It 'leaves an unfinished attempt when a thread update fails, without failing the script' {
        $stateDir = New-StateDir
        $environment = $CommentEnvironment.Clone()
        $environment.COPILOT_REVIEW_STATE_DIR = $stateDir
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 500; Body = '{"message":"internal server error"}' }
        ) -Files (Get-PolicyFiles $stateDir) -Environment $environment -BuildArgs {
            param($baseUrl)
            @('-File', (Join-Path $ScriptsDir 'Update-CopilotComment.ps1'), '-ThreadId', '42', '-Status', 'Fixed')
        }

        # Updates stay non-blocking for the agent...
        $result.ExitCode | Should -Be 0
        # ...but the task still learns the attempt did not complete
        (Get-ActivitySummary $stateDir).Unfinished | Should -Be 1
    }

    It 'treats an invalid update request as an unfinished attempt: <Case>' -ForEach @(
        @{ Case = 'invalid status'; Arguments = @('-ThreadId', '42', '-Status', 'Resolved') },
        @{ Case = 'zero thread ID'; Arguments = @('-ThreadId', '0', '-Status', 'Fixed') },
        @{ Case = 'missing thread ID'; Arguments = @('-Status', 'Fixed') },
        @{ Case = 'non-numeric comment ID'; Arguments = @('-ThreadId', '42', '-CommentId', 'abc', '-Content', 'Updated') }
    ) {
        $stateDir = New-StateDir
        $environment = $CommentEnvironment.Clone()
        $environment.COPILOT_REVIEW_STATE_DIR = $stateDir
        $updateArgs = $Arguments
        $updateScript = Join-Path $ScriptsDir 'Update-CopilotComment.ps1'
        $result = Invoke-ScriptWithMockApi -Responses @() -Files (Get-PolicyFiles $stateDir) -Environment $environment -BuildArgs {
            param($baseUrl)
            @('-File', $updateScript) + $updateArgs
        }.GetNewClosure()

        # Updates stay non-blocking (exit 0), but the attempt never finishes
        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 0
        (Get-ActivitySummary $stateDir).Unfinished | Should -Be 1
    }

    It 'writes nothing when the task has not set a state folder' {
        # An empty variable must not turn into a relative 'activity' folder in
        # the working directory (the agent's checkout)
        $workDir = New-StateDir
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $ThreadJson }
        ) -Environment @{ COPILOT_REVIEW_STATE_DIR = '' } -WorkingDirectory $workDir -BuildArgs {
            param($baseUrl)
            @('-File', (Join-Path $ScriptsDir 'Add-AzureDevOpsPRComment.ps1'),
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '123', '-Comment', 'Feedback')
        }
        $result.ExitCode | Should -Be 0
        @(Get-ChildItem -Path $workDir -Force).Count | Should -Be 0
    }
}

Describe 'Context degradation markers' {

    BeforeAll {
        $script:PrDetailsJson = @{
            pullRequestId = 7; title = 'PR'; status = 'active'; repository = @{ name = 'TestRepo' }
            sourceRefName = 'refs/heads/feature'; targetRefName = 'refs/heads/main'; isDraft = $false
            mergeStatus = 'succeeded'; createdBy = @{ displayName = 'Dev'; uniqueName = 'dev@example.com' }
            creationDate = '2026-01-01T00:00:00Z'; reviewers = @()
        } | ConvertTo-Json -Depth 4

        function Invoke-PrDetailsWithThreads {
            param([hashtable]$ThreadsResponse, [string]$StateDir)
            $outFile = Join-Path $StateDir 'PR_Details.txt'
            $prScript = Join-Path $ScriptsDir 'Get-AzureDevOpsPR.ps1'
            Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 200; Body = $PrDetailsJson },
                @{ Status = 200; Body = '{"count":0,"value":[]}' },
                @{ Status = 200; Body = '{"count":1,"value":[{"id":1}]}' },
                $ThreadsResponse
            ) -Environment @{ COPILOT_REVIEW_STATE_DIR = $StateDir } -BuildArgs {
                param($baseUrl)
                @('-File', $prScript, '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
                  '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '7', '-OutputFile', $outFile)
            }.GetNewClosure()
        }
    }

    It 'flags failed thread retrieval instead of reporting no comments' {
        $stateDir = New-StateDir
        $result = Invoke-PrDetailsWithThreads -StateDir $stateDir -ThreadsResponse @{ Status = 500; Body = '{"message":"boom"}' }

        $result.ExitCode | Should -Be 0
        Test-Path (Join-Path $stateDir 'context/pr-threads.degraded') | Should -BeTrue
        $text = Get-Content -Path (Join-Path $stateDir 'PR_Details.txt') -Raw
        $text | Should -Match 'Existing comment threads could not be retrieved'
        $text | Should -Not -Match 'No comments'
    }

    It 'writes no marker when everything loads' {
        $stateDir = New-StateDir
        $result = Invoke-PrDetailsWithThreads -StateDir $stateDir -ThreadsResponse @{ Status = 200; Body = '{"count":0,"value":[]}' }

        $result.ExitCode | Should -Be 0
        Test-Path (Join-Path $stateDir 'context') | Should -BeFalse
        Get-Content -Path (Join-Path $stateDir 'PR_Details.txt') -Raw | Should -Match 'No comments'
    }

    It 'flags a failed work item batch fetch, which still exits 0' {
        $stateDir = New-StateDir
        $outFile = Join-Path $stateDir 'Work_Item_Details.txt'
        $workItemsScript = Join-Path $ScriptsDir 'Get-AzureDevOpsWorkItems.ps1'
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 500; Body = '{"message":"boom"}' }
        ) -Environment @{ COPILOT_REVIEW_STATE_DIR = $stateDir } -BuildArgs {
            param($baseUrl)
            @('-File', $workItemsScript, '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-WorkItemIds', '5,6', '-OutputFile', $outFile)
        }.GetNewClosure()

        $result.ExitCode | Should -Be 0
        Test-Path (Join-Path $stateDir 'context/work-item-details.degraded') | Should -BeTrue
    }
}
