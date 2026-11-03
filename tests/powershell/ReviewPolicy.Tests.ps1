<#
    Tests for the review policy the comment scripts enforce (PR #65 rework).

    Before the review agent starts, the task writes review-policy.json to its
    per-run state folder. The comment scripts check every action against it:
    the destination (collection URI, project, repository, PR id), the threads
    the agent may reply to or resolve, the comments it may edit or delete,
    where comment text may be read from, and the files inline comments may
    anchor to. A secret check applies to everything they send, also outside the
    task. A refusal must happen before any request, exit 1, leave the attempt
    unfinished (so the run is not recorded as reviewed), and leave a refusal
    record the task turns into a pipeline warning.

    Scripts run as real child pwsh processes against a local HttpListener
    (MockApiHelper.ps1). Requires Pester 5+. Run via tests/powershell/Invoke-Tests.ps1.
#>

BeforeDiscovery {
    $script:PolicyScripts = @('Add-CopilotComment.ps1', 'Add-AzureDevOpsPRComment.ps1', 'Update-CopilotComment.ps1', 'Delete-CopilotComment.ps1')
}

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:ScriptsDir = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts'
    . (Join-Path $PSScriptRoot 'MockApiHelper.ps1')

    # The shared policy region of a script (line endings normalized)
    function Get-PolicyRegionText {
        param([string]$ScriptName)
        $text = [System.IO.File]::ReadAllText((Join-Path $ScriptsDir $ScriptName)) -replace "`r`n", "`n"
        $match = [regex]::Match($text, '(?s)#region Review policy\n.*?#endregion Review policy')
        if ($match.Success) { $match.Value } else { $null }
    }

    function New-TestDir {
        $dir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $dir | Out-Null
        $dir
    }

    # A fresh state folder and working directory, as the task creates per run
    function New-PolicyRun {
        [pscustomobject]@{ StateDir = (New-TestDir); WorkDir = (New-TestDir) }
    }

    $script:CommentEnvironment = @{
        AZUREDEVOPS_TOKEN          = 'test-token'
        AZUREDEVOPS_AUTH_TYPE      = 'Basic'
        AZUREDEVOPS_COLLECTION_URI = '__BASEURL__/testorg'
        PROJECT                    = 'TestProject'
        REPOSITORY                 = 'TestRepo'
        PRID                       = '123'
    }

    # Runs a comment script the way the review agent does during a run: state
    # folder set (with the given policy JSON, or none when $null), working
    # directory = the checkout. -Standalone leaves the state folder unset.
    # Credential variables inherited from the developer's shell are cleared so
    # the secret check sees only what a test sets.
    function Invoke-PolicyScript {
        param(
            [Parameter(Mandatory)] [string]$Script,
            [Parameter(Mandatory)] $Run,
            $Policy,
            [string[]]$Arguments = @(),
            [hashtable[]]$Responses = @(),
            [hashtable]$Environment = @{},
            [switch]$Standalone
        )

        $childEnvironment = $CommentEnvironment.Clone()
        foreach ($name in @('SYSTEM_ACCESSTOKEN', 'ANTHROPIC_API_KEY', 'CLAUDE_CODE_OAUTH_TOKEN', 'GH_TOKEN', 'GITHUB_TOKEN',
                'COPILOT_GITHUB_TOKEN', 'COPILOT_PROVIDER_API_KEY', 'COPILOT_PROVIDER_BEARER_TOKEN')) {
            $childEnvironment[$name] = ''
        }
        foreach ($key in $Environment.Keys) {
            $childEnvironment[$key] = $Environment[$key]
        }
        $files = @{}
        if ($Standalone) {
            $childEnvironment['COPILOT_REVIEW_STATE_DIR'] = ''
        }
        else {
            $childEnvironment['COPILOT_REVIEW_STATE_DIR'] = $Run.StateDir
            if ($null -ne $Policy) {
                $files[(Join-Path $Run.StateDir 'review-policy.json')] = [string]$Policy
            }
        }

        # GetNewClosure captures only locals
        $scriptPath = Join-Path $ScriptsDir $Script
        $scriptArgs = $Arguments
        Invoke-ScriptWithMockApi -Responses $Responses -Environment $childEnvironment -Files $files `
            -WorkingDirectory $Run.WorkDir -BuildArgs {
            param($baseUrl)
            @('-File', $scriptPath) + @($scriptArgs | ForEach-Object { $_ -replace '__BASEURL__', $baseUrl })
        }.GetNewClosure()
    }

    # Arguments for calling Add-AzureDevOpsPRComment.ps1 directly
    function Get-AddArgs {
        param(
            [string[]]$Extra = @(),
            [string]$CollectionUri = '__BASEURL__/testorg',
            [string]$Project = 'TestProject',
            [string]$Repository = 'TestRepo',
            [string]$Id = '123',
            [string]$Token = 'test-token',
            [string]$Comment = 'Feedback'
        )
        @('-Token', $Token, '-CollectionUri', $CollectionUri, '-Project', $Project, '-Repository', $Repository,
          '-Id', $Id, '-Comment', $Comment) + $Extra
    }

    # Activity markers: unfinished attempts, finished actions, refusal reasons
    function Get-PolicyActivity {
        param([string]$StateDir)
        $dir = Join-Path $StateDir 'activity'
        $files = if (Test-Path $dir) { @(Get-ChildItem -LiteralPath $dir -File) } else { @() }
        $started = @($files | Where-Object { $_.Extension -eq '.started' })
        [pscustomobject]@{
            Unfinished = @($started | Where-Object { -not (Test-Path ([System.IO.Path]::ChangeExtension($_.FullName, '.done'))) }).Count
            Done       = @($files | Where-Object { $_.Extension -eq '.done' } | ForEach-Object { (Get-Content -LiteralPath $_.FullName -Raw).Trim() } | Sort-Object)
            Refusals   = @($files | Where-Object { $_.Extension -eq '.refused' } | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw })
        }
    }

    # A refusal: nothing sent, exit 1, reason on stderr, the attempt left
    # unfinished, and exactly one refusal record for the task
    function Assert-Refused {
        param($Result, $Run, [string]$Pattern)
        $Result.ExitCode | Should -Be 1
        $Result.Requests.Count | Should -Be 0
        $Result.Errors | Should -Match 'Refused by review policy'
        if ($Pattern) {
            $Result.Errors | Should -Match $Pattern
        }
        $activity = Get-PolicyActivity $Run.StateDir
        $activity.Unfinished | Should -BeGreaterThan 0
        $activity.Refusals.Count | Should -Be 1
    }

    $script:PrJson = '{"title":"Test PR"}'
    $script:CreatedThreadJson = '{"id":77,"status":"active","comments":[{"id":1,"author":{"displayName":"CI Mock"},"publishedDate":"2026-01-01T00:00:00Z"}]}'
    $script:ExistingThreadJson = '{"id":42,"status":"active","comments":[{"id":1,"content":"Finding","author":{"displayName":"CI Mock"}}]}'
    $script:ReplyJson = '{"id":2,"author":{"displayName":"CI Mock"},"publishedDate":"2026-01-01T00:00:00Z"}'
    $script:Secret = 'sk-test-0123456789abcdef'
}

Describe 'Review policy region' {

    It 'is identical in <_>' -ForEach $PolicyScripts {
        $text = Get-PolicyRegionText $_
        $text | Should -Not -BeNullOrEmpty
        $text | Should -BeExactly (Get-PolicyRegionText 'Add-AzureDevOpsPRComment.ps1')
    }
}

Describe 'Get-AzureDevOpsPR.ps1 thread data for the review policy' {

    BeforeAll {
        $comment = {
            param([int]$Id, [string]$Content, [string]$Author, [string]$Type = 'text')
            @{ id = $Id; commentType = $Type; content = $Content; publishedDate = '2026-01-02T00:00:00Z'
               author = @{ displayName = $Author; uniqueName = "$($Author -replace ' ', '.')@example.com" } }
        }
        $script:PolicyThreadsJson = @{
            count = 5
            value = @(
                # Automated-review thread with a human reply and a tagged reply posted with a PAT
                @{ id = 31; status = 'active'; comments = @(
                    (& $comment 1 'Agent finding' 'TestProject Build Service (testorg)'),
                    (& $comment 2 'I disagree' 'Dev'),
                    (& $comment 3 'Fixed now. _[Generated by GitHub Copilot]_' 'Pat Owner')) },
                # Human thread the agent replied to
                @{ id = 32; status = 'active'; comments = @(
                    (& $comment 1 'Human concern' 'Jane Reviewer'),
                    (& $comment 2 'Addressed. _[Generated by Claude Code]_' 'Pat Owner')) },
                # System thread (e.g. a push notification): never actionable
                @{ id = 33; status = 'active'; comments = @((& $comment 1 'Dev pushed 1 commit' 'Dev' 'system')) },
                # Human thread with a system message inside
                @{ id = 34; status = 'fixed'; comments = @(
                    (& $comment 1 'Typo here' 'Jane Reviewer'),
                    (& $comment 2 'Status changed' 'Jane Reviewer' 'system')) }
            )
        } | ConvertTo-Json -Depth 8

        $script:PrDetailsJson = @{
            pullRequestId = 9; title = 'PR'; status = 'active'; repository = @{ name = 'TestRepo' }
            sourceRefName = 'refs/heads/feature'; targetRefName = 'refs/heads/main'; isDraft = $false
            mergeStatus = 'succeeded'; createdBy = @{ displayName = 'Dev'; uniqueName = 'dev@example.com' }
            creationDate = '2026-01-01T00:00:00Z'; reviewers = @()
        } | ConvertTo-Json -Depth 4

        function Invoke-ThreadPolicy {
            param([hashtable]$ThreadsResponse)
            $outDir = New-TestDir
            $policyFile = Join-Path $outDir 'PR_Threads.json'
            $prScript = Join-Path $ScriptsDir 'Get-AzureDevOpsPR.ps1'
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 200; Body = $PrDetailsJson },
                @{ Status = 200; Body = '{"count":0,"value":[]}' },
                @{ Status = 200; Body = '{"count":1,"value":[{"id":1}]}' },
                $ThreadsResponse
            ) -BuildArgs {
                param($baseUrl)
                @('-File', $prScript, '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg", '-Project', 'TestProject',
                  '-Repository', 'TestRepo', '-Id', '9', '-OutputFile', (Join-Path $outDir 'PR_Details.txt'),
                  '-ThreadPolicyFile', $policyFile)
            }.GetNewClosure()
            $result | Add-Member -NotePropertyName PolicyText -NotePropertyValue $(if (Test-Path $policyFile) { Get-Content -LiteralPath $policyFile -Raw } else { $null }) -PassThru
        }

        $script:ThreadRun = Invoke-ThreadPolicy -ThreadsResponse @{ Status = 200; Body = $PolicyThreadsJson }
        $script:ThreadData = $ThreadRun.PolicyText | ConvertFrom-Json
    }

    It 'lists the comment threads with their authorship, leaving out system threads' {
        $ThreadRun.ExitCode | Should -Be 0
        $ThreadData.threadsStatus | Should -Be 'Complete'
        @($ThreadData.threads.threadId) | Should -Be @(31, 32, 34)
        @($ThreadData.threads | ForEach-Object { $_.agentAuthored }) | Should -Be @($true, $false, $false)
        @($ThreadData.threads | ForEach-Object { $_.status }) | Should -Be @('active', 'active', 'fixed')
    }

    It 'marks each comment by whether an automated review wrote it, leaving out system messages' {
        $agentThread = $ThreadData.threads | Where-Object threadId -eq 31
        @($agentThread.comments | ForEach-Object { "$($_.commentId)=$($_.agentAuthored)" }) | Should -Be @('1=True', '2=False', '3=True')
        $humanThread = $ThreadData.threads | Where-Object threadId -eq 32
        @($humanThread.comments | ForEach-Object { "$($_.commentId)=$($_.agentAuthored)" }) | Should -Be @('1=False', '2=True')
        $withSystem = $ThreadData.threads | Where-Object threadId -eq 34
        @($withSystem.comments.commentId) | Should -Be @(1)
        # Arrays stay arrays in the JSON, even with one element
        $ThreadRun.PolicyText | Should -Match '"comments":\s*\[\s*\{\s*"commentId":\s*1,'
    }

    It 'reports failed thread retrieval instead of an empty list' {
        $failed = Invoke-ThreadPolicy -ThreadsResponse @{ Status = 500; Body = '{"message":"boom"}' }
        $failed.ExitCode | Should -Be 0
        $data = $failed.PolicyText | ConvertFrom-Json
        $data.threadsStatus | Should -Be 'Failed'
        @($data.threads).Count | Should -Be 0
    }
}

Describe 'Add-AzureDevOpsPRComment.ps1 under a review policy' {

    It 'posts a general comment to the configured pull request' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Arguments (Get-AddArgs) -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 2
        $activity = Get-PolicyActivity $run.StateDir
        $activity.Unfinished | Should -Be 0
        $activity.Done | Should -Be @('thread')
        $activity.Refusals.Count | Should -Be 0
    }

    It 'accepts the configured destination written differently (case, trailing slash)' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Arguments (Get-AddArgs -CollectionUri '__BASEURL__/TestOrg/' -Project 'testproject' -Repository 'TESTREPO') `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 2
    }

    It 'refuses another <Name> before making any request' -ForEach @(
        @{ Name = 'collection'; Overrides = @{ CollectionUri = '__BASEURL__/otherorg' } },
        @{ Name = 'server'; Overrides = @{ CollectionUri = 'http://127.0.0.2:9/testorg' } },
        @{ Name = 'collection URI with a query'; Overrides = @{ CollectionUri = '__BASEURL__/testorg?x=1' } },
        @{ Name = 'project'; Overrides = @{ Project = 'OtherProject' } },
        @{ Name = 'repository'; Overrides = @{ Repository = 'OtherRepo' } },
        @{ Name = 'pull request'; Overrides = @{ Id = '124' } }
    ) {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Arguments (Get-AddArgs @Overrides)

        Assert-Refused $result $run 'may act only on pull request #123'
    }

    It 'refuses a reply to a thread the review may not act on' {
        $run = New-PolicyRun
        $policy = New-TestPolicyJson -WorkspaceRoot $run.WorkDir -Threads @(New-TestPolicyThread -ThreadId 42 -MayAct $false -AgentAuthored $false)
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy $policy -Arguments (Get-AddArgs -Extra @('-ThreadId', '42'))

        Assert-Refused $result $run 'may not reply to thread #42'
    }

    It 'refuses a reply to a thread the policy does not list' {
        $run = New-PolicyRun
        $policy = New-TestPolicyJson -WorkspaceRoot $run.WorkDir -Threads @(New-TestPolicyThread -ThreadId 42)
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy $policy -Arguments (Get-AddArgs -Extra @('-ThreadId', '99'))

        Assert-Refused $result $run 'thread #99'
    }

    It 'replies to a human thread the policy allows' {
        # e.g. resolveThreads 'all' with an active human thread, or a raw prompt
        $run = New-PolicyRun
        $policy = New-TestPolicyJson -WorkspaceRoot $run.WorkDir -Threads @(New-TestPolicyThread -ThreadId 42 -AgentAuthored $false)
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy $policy -Arguments (Get-AddArgs -Extra @('-ThreadId', '42')) `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $ExistingThreadJson }, @{ Status = 200; Body = $ReplyJson })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 3
        $result.Requests[2].Path | Should -Match '/threads/42/comments'
        (Get-PolicyActivity $run.StateDir).Done | Should -Be @('reply')
    }

    It 'refuses replies when the threads could not be loaded' {
        $run = New-PolicyRun
        $policy = New-TestPolicyJson -WorkspaceRoot $run.WorkDir -ThreadsComplete $false
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy $policy -Arguments (Get-AddArgs -Extra @('-ThreadId', '42'))

        Assert-Refused $result $run 'could not be loaded'
    }

    It 'anchors an inline comment to a file the pull request changes' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir -ChangedFiles @('/src/App.cs')) `
            -Arguments (Get-AddArgs -Extra @('-FilePath', 'src\App.cs', '-StartLine', '10')) `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        $posted = $result.Requests[1].Body | ConvertFrom-Json
        $posted.threadContext.filePath | Should -Be '/src/App.cs'
        $posted.comments[0].content | Should -BeExactly 'Feedback'
    }

    It 'posts a comment on a file the pull request does not change as a general comment' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir -ChangedFiles @('/src/App.cs')) `
            -Arguments (Get-AddArgs -Extra @('-FilePath', '/src/Other.cs', '-StartLine', '10', '-EndLine', '12')) `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 2
        $result.Output | Should -Match 'is not a file changed in this pull request'
        $posted = $result.Requests[1].Body | ConvertFrom-Json
        $posted.PSObject.Properties.Name | Should -Not -Contain 'threadContext'
        $posted.comments[0].content | Should -BeExactly "Feedback`n`n**File:** ``/src/Other.cs```n**Lines 10-12**"
        # Not a refusal: the attempt finishes and the run can be recorded
        $activity = Get-PolicyActivity $run.StateDir
        $activity.Unfinished | Should -Be 0
        $activity.Refusals.Count | Should -Be 0
    }

    It 'with an incomplete change list, anchors to a file that exists in the working directory' {
        $run = New-PolicyRun
        New-Item -ItemType Directory -Path (Join-Path $run.WorkDir 'src') | Out-Null
        Set-Content -LiteralPath (Join-Path $run.WorkDir 'src/Exists.cs') -Value 'class C {}'
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run `
            -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir -ChangedFiles @() -ChangedFilesComplete $false) `
            -Arguments (Get-AddArgs -Extra @('-FilePath', '/src/Exists.cs', '-StartLine', '1')) `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        ($result.Requests[1].Body | ConvertFrom-Json).threadContext.filePath | Should -Be '/src/Exists.cs'
    }

    It 'with an incomplete change list, posts a comment on a missing file as a general comment' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run `
            -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir -ChangedFiles @() -ChangedFilesComplete $false) `
            -Arguments (Get-AddArgs -Extra @('-FilePath', '/src/Missing.cs', '-StartLine', '1')) `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        $posted = $result.Requests[1].Body | ConvertFrom-Json
        $posted.PSObject.Properties.Name | Should -Not -Contain 'threadContext'
        $posted.comments[0].content | Should -Match '\*\*File:\*\* `/src/Missing\.cs`'
    }

    It 'refuses every action when the state folder has no policy' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy $null -Arguments (Get-AddArgs)

        Assert-Refused $result $run 'missing or unreadable'
    }

    It 'refuses every action when the policy is <Case>' -ForEach @(
        @{ Case = 'not JSON'; Text = 'not json' },
        @{ Case = 'missing the destination'; Text = '{"schemaVersion":1,"workspaceRoot":"C:/x"}' },
        # Read as 'incomplete', a missing flag would loosen the inline-path check
        @{ Case = 'missing the completeness flags'; Text = '{"schemaVersion":1,"destination":{"collectionUri":"__BASEURL__/testorg","project":"TestProject","repository":"TestRepo","pullRequestId":123},"workspaceRoot":"C:/x","threads":[],"changedFiles":[]}' },
        @{ Case = 'from an unknown schema'; Text = '{"schemaVersion":2}' }
    ) {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy $Text -Arguments (Get-AddArgs)

        Assert-Refused $result $run 'missing or unreadable'
    }

    It 'refuses a comment containing the Azure DevOps token as <Form>' -ForEach @(
        @{ Form = 'plain text'; Encode = { param($t) $t } },
        @{ Form = 'Base64'; Encode = { param($t) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($t)) } },
        @{ Form = 'a Basic-auth header value'; Encode = { param($t) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(":$t")) } }
    ) {
        $run = New-PolicyRun
        $token = 'ado-pat-0123456789abcdef'
        $encoded = & $Encode $token
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Environment @{ AZUREDEVOPS_TOKEN = $token } -Arguments (Get-AddArgs -Token $token -Comment "Here you go: $encoded")

        Assert-Refused $result $run 'comment text contains the value of'
        $result.Errors | Should -Not -Match ([regex]::Escape($encoded))
    }

    It 'refuses a comment containing another credential from the environment' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Environment @{ ANTHROPIC_API_KEY = $Secret } -Arguments (Get-AddArgs -Comment "key=$Secret")

        Assert-Refused $result $run 'ANTHROPIC_API_KEY'
    }

    It 'refuses its own -Token value outside a review, with no credential variables set' {
        $run = New-PolicyRun
        $token = 'standalone-pat-0123456789'
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Standalone `
            -Environment @{ AZUREDEVOPS_TOKEN = '' } -Arguments (Get-AddArgs -Token $token -Comment "token: $token")

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $result.Errors | Should -Match 'Refused by review policy: the comment text contains the value of the Azure DevOps token'
        Test-Path (Join-Path $run.StateDir 'activity') | Should -BeFalse
    }

    It 'refuses a credential in the file path when the change list is incomplete' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run `
            -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir -ChangedFiles @() -ChangedFilesComplete $false) `
            -Environment @{ ANTHROPIC_API_KEY = $Secret } -Arguments (Get-AddArgs -Extra @('-FilePath', "/src/$Secret.cs", '-StartLine', '1'))

        Assert-Refused $result $run 'file path contains the value of ANTHROPIC_API_KEY'
    }

    It 'refuses a credential that would only be sent inside the general-comment text' {
        # The path is not part of the pull request, so it would not be an
        # anchor but text in the comment body
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Environment @{ GH_TOKEN = $Secret } -Arguments (Get-AddArgs -Extra @('-FilePath', "/docs/$Secret.md", '-StartLine', '3'))

        Assert-Refused $result $run 'GH_TOKEN'
    }

    It 'prints nothing derived from a credential-bearing path before refusing it' {
        # The off-PR note would otherwise echo the path before the secret check
        $run = New-PolicyRun
        $token = 'ado-pat-0123456789abcdef'
        $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($token))
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Environment @{ AZUREDEVOPS_TOKEN = $token } -Arguments (Get-AddArgs -Token $token -Extra @('-FilePath', "/src/$encoded.cs", '-StartLine', '1'))

        Assert-Refused $result $run 'file path contains the value of'
        $result.Output | Should -Not -Match ([regex]::Escape($encoded))
        $result.Errors | Should -Not -Match ([regex]::Escape($encoded))
    }

    It 'masks credentials in the refusal reason' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Add-AzureDevOpsPRComment.ps1' -Run $run -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Environment @{ ANTHROPIC_API_KEY = $Secret } -Arguments (Get-AddArgs -Project "Proj-$Secret")

        Assert-Refused $result $run 'Proj-\*\*\*'
        $result.Errors | Should -Not -Match ([regex]::Escape($Secret))
        (Get-PolicyActivity $run.StateDir).Refusals[0] | Should -Not -Match ([regex]::Escape($Secret))
    }
}

Describe 'Add-CopilotComment.ps1 under a review policy' {

    BeforeAll {
        function Invoke-Wrapper {
            param($Run, [string[]]$Arguments, [hashtable[]]$Responses = @(), $Policy)
            if (-not $PSBoundParameters.ContainsKey('Policy')) {
                $Policy = New-TestPolicyJson -WorkspaceRoot $Run.WorkDir -Threads @(New-TestPolicyThread -ThreadId 42)
            }
            Invoke-PolicyScript -Script 'Add-CopilotComment.ps1' -Run $Run -Policy $Policy -Arguments $Arguments -Responses $Responses
        }
    }

    It 'refuses inline -Comment text' {
        $run = New-PolicyRun
        $result = Invoke-Wrapper -Run $run -Arguments @('-Comment', 'Inline text')

        Assert-Refused $result $run 'pass -CommentFile'
    }

    It 'posts the text of a -CommentFile in the working directory' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_comment.md') -Value 'From a file' -NoNewline
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './_comment.md', '-Status', 'Closed') `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        ($result.Requests[1].Body | ConvertFrom-Json).comments[0].content | Should -BeExactly 'From a file'
        $activity = Get-PolicyActivity $run.StateDir
        $activity.Unfinished | Should -Be 0
        $activity.Done | Should -Be @('thread', 'wrapper')
    }

    It 'replies from a -CommentFile to a thread the review may act on' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_reply.md') -Value 'Addressed.'
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './_reply.md', '-ThreadId', '42') `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $ExistingThreadJson }, @{ Status = 200; Body = $ReplyJson })

        $result.ExitCode | Should -Be 0
        (Get-PolicyActivity $run.StateDir).Done | Should -Be @('reply', 'wrapper')
    }

    It 'refuses a -CommentFile <Case>' -ForEach @(
        @{ Case = 'outside the working directory'; Kind = 'outside' },
        @{ Case = 'reached with ..'; Kind = 'dotdot' },
        @{ Case = 'inside the .git folder'; Kind = 'git' }
    ) {
        $run = New-PolicyRun
        $outsideDir = New-TestDir
        Set-Content -LiteralPath (Join-Path $outsideDir 'secret.txt') -Value 'not for posting'
        New-Item -ItemType Directory -Path (Join-Path $run.WorkDir '.git') | Out-Null
        Set-Content -LiteralPath (Join-Path $run.WorkDir '.git/config') -Value '[http] extraheader = AUTHORIZATION: bearer x'
        $path = switch ($Kind) {
            'outside' { Join-Path $outsideDir 'secret.txt' }
            'dotdot' { "../$(Split-Path -Leaf $outsideDir)/secret.txt" }
            'git' { './.git/config' }
        }
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', $path)

        Assert-Refused $result $run $(if ($Kind -eq 'git') { 'inside a \.git folder' } else { 'outside the working directory' })
    }

    It 'refuses a -CommentFile that is a link to a file outside the working directory' {
        $run = New-PolicyRun
        $outsideDir = New-TestDir
        Set-Content -LiteralPath (Join-Path $outsideDir 'secret.txt') -Value 'not for posting'
        try {
            New-Item -ItemType SymbolicLink -Path (Join-Path $run.WorkDir '_comment.md') -Target (Join-Path $outsideDir 'secret.txt') -ErrorAction Stop | Out-Null
        }
        catch {
            Set-ItResult -Skipped -Because "symbolic links cannot be created here: $($_.Exception.Message)"
            return
        }
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './_comment.md')

        Assert-Refused $result $run 'outside the working directory'
    }

    It 'refuses a -CommentFile inside a folder link that leads outside the working directory' {
        $run = New-PolicyRun
        $outsideDir = New-TestDir
        Set-Content -LiteralPath (Join-Path $outsideDir 'secret.txt') -Value 'not for posting'
        # A junction needs no special rights on Windows; elsewhere a directory symlink
        $linkType = if ($IsWindows) { 'Junction' } else { 'SymbolicLink' }
        New-Item -ItemType $linkType -Path (Join-Path $run.WorkDir 'docs') -Target $outsideDir | Out-Null
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './docs/secret.txt')

        Assert-Refused $result $run 'outside the working directory'
    }

    It 'refuses a -CommentFile with more than 100,000 characters' {
        $run = New-PolicyRun
        [System.IO.File]::WriteAllText((Join-Path $run.WorkDir 'long.md'), ('x' * 100001))
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './long.md')

        Assert-Refused $result $run 'at most 100,000'
    }

    It 'accepts a -CommentFile of exactly 100,000 characters' {
        $run = New-PolicyRun
        [System.IO.File]::WriteAllText((Join-Path $run.WorkDir 'long.md'), ('x' * 100000))
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './long.md') `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        ($result.Requests[1].Body | ConvertFrom-Json).comments[0].content.Length | Should -Be 100000
    }

    It 'refuses a credential in the file path before printing it' {
        $run = New-PolicyRun
        $token = 'ado-pat-0123456789abcdef'
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_comment.md') -Value 'Finding'
        $environment = @{ AZUREDEVOPS_TOKEN = $token }
        $result = Invoke-PolicyScript -Script 'Add-CopilotComment.ps1' -Run $run -Environment $environment `
            -Policy (New-TestPolicyJson -WorkspaceRoot $run.WorkDir) `
            -Arguments @('-CommentFile', './_comment.md', '-FilePath', "/src/$token.cs", '-StartLine', '1')

        Assert-Refused $result $run 'file path contains the value of'
        $result.Output | Should -Not -Match ([regex]::Escape($token))
        # Refused by the wrapper itself: the posting script never started
        (Get-PolicyActivity $run.StateDir).Unfinished | Should -Be 1
    }

    It 'reports a missing -CommentFile as an error, not a refusal' {
        $run = New-PolicyRun
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './missing.md')

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $result.Errors | Should -Match 'Comment file not found'
        $activity = Get-PolicyActivity $run.StateDir
        $activity.Unfinished | Should -Be 1
        $activity.Refusals.Count | Should -Be 0
    }

    It 'refuses a reply to a thread the review may not act on, through the posting script' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_reply.md') -Value 'Resolved.'
        $result = Invoke-Wrapper -Run $run -Arguments @('-CommentFile', './_reply.md', '-ThreadId', '55')

        Assert-Refused $result $run 'thread #55'
        # Both the wrapper and the posting script it called stay unfinished
        (Get-PolicyActivity $run.StateDir).Unfinished | Should -Be 2
    }
}

Describe 'Prompt template commands under a review policy' {

    BeforeAll {
        # Runs a command line exactly as the prompt templates give it to the
        # agent: the comment scripts copied into the working directory (as the
        # task does) and dot-sourced through pwsh -Command
        function Invoke-TemplateCommand {
            param($Run, [string]$Command, [hashtable[]]$Responses)
            foreach ($name in 'Add-CopilotComment.ps1', 'Add-AzureDevOpsPRComment.ps1', 'Update-CopilotComment.ps1') {
                Copy-Item -LiteralPath (Join-Path $ScriptsDir $name) -Destination $Run.WorkDir
            }
            $environment = $CommentEnvironment.Clone()
            $environment.COPILOT_REVIEW_STATE_DIR = $Run.StateDir
            $policy = New-TestPolicyJson -WorkspaceRoot $Run.WorkDir -Threads @(New-TestPolicyThread -ThreadId 42)
            $commandLine = $Command
            Invoke-ScriptWithMockApi -Responses $Responses -Environment $environment -WorkingDirectory $Run.WorkDir `
                -Files @{ (Join-Path $Run.StateDir 'review-policy.json') = $policy } -BuildArgs {
                param($baseUrl)
                @('-Command', $commandLine)
            }.GetNewClosure()
        }
    }

    It 'replies from _reply.md the way the templates instruct' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_reply.md') -Value 'Issue addressed in the current iteration. Marking as resolved. _[Generated by GitHub Copilot]_'
        $result = Invoke-TemplateCommand -Run $run -Command ". ./Add-CopilotComment.ps1 -ThreadId 42 -CommentFile './_reply.md'" `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $ExistingThreadJson }, @{ Status = 200; Body = $ReplyJson })

        $result.ExitCode | Should -Be 0
        ($result.Requests[2].Body | ConvertFrom-Json).content | Should -Match 'Marking as resolved'
        (Get-PolicyActivity $run.StateDir).Done | Should -Be @('reply', 'wrapper')
    }

    It 'posts an inline comment from _comment.md the way the templates instruct' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_comment.md') -Value 'Finding. _[Generated by GitHub Copilot]_'
        $result = Invoke-TemplateCommand -Run $run -Command ". ./Add-CopilotComment.ps1 -CommentFile './_comment.md' -Status 'Active' -FilePath '/src/App.cs' -StartLine 42 -EndLine 45" `
            -Responses @(@{ Status = 200; Body = $PrJson }, @{ Status = 200; Body = $CreatedThreadJson })

        $result.ExitCode | Should -Be 0
        ($result.Requests[1].Body | ConvertFrom-Json).threadContext.filePath | Should -Be '/src/App.cs'
    }

    It 'resolves a thread the way the templates instruct' {
        $run = New-PolicyRun
        $result = Invoke-TemplateCommand -Run $run -Command '. ./Update-CopilotComment.ps1 -ThreadId 42 -Status Fixed' `
            -Responses @(@{ Status = 200; Body = '{"id":42,"status":"fixed"}' })

        $result.ExitCode | Should -Be 0
        $result.Requests[0].Method | Should -Be 'PATCH'
        (Get-PolicyActivity $run.StateDir).Done | Should -Be @('update')
    }

    It 'refuses the old inline reply form' {
        $run = New-PolicyRun
        $result = Invoke-TemplateCommand -Run $run -Responses @() `
            -Command ". ./Add-CopilotComment.ps1 -ThreadId 42 -Comment 'Issue addressed in the current iteration. Marking as resolved. _[Generated by GitHub Copilot]_'"

        Assert-Refused $result $run 'pass -CommentFile'
    }
}

Describe 'Update-CopilotComment.ps1 under a review policy' {

    BeforeAll {
        function Invoke-Update {
            param($Run, [string[]]$Arguments, [hashtable[]]$Responses = @(), [hashtable]$Environment = @{})
            $policy = New-TestPolicyJson -WorkspaceRoot $Run.WorkDir -Threads @(
                (New-TestPolicyThread -ThreadId 42 -AgentCommentIds 1),
                (New-TestPolicyThread -ThreadId 43 -MayAct $false -AgentAuthored $false -AgentCommentIds 3)
            )
            Invoke-PolicyScript -Script 'Update-CopilotComment.ps1' -Run $Run -Policy $policy -Arguments $Arguments `
                -Responses $Responses -Environment $Environment
        }
    }

    It 'resolves a thread the review may act on' {
        $run = New-PolicyRun
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-Status', 'Fixed') -Responses @(@{ Status = 200; Body = '{"id":42,"status":"fixed"}' })

        $result.ExitCode | Should -Be 0
        $result.Requests[0].Method | Should -Be 'PATCH'
        (Get-PolicyActivity $run.StateDir).Done | Should -Be @('update')
    }

    It 'refuses to change the status of a thread the review may not act on' {
        $run = New-PolicyRun
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '43', '-Status', 'Fixed')

        Assert-Refused $result $run 'may not change the status of thread #43'
    }

    It 'refuses inline -Content' {
        $run = New-PolicyRun
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '1', '-Content', 'Edited')

        Assert-Refused $result $run 'pass -ContentFile'
    }

    It 'edits a comment an automated review wrote, from a -ContentFile' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_update.md') -Value 'Edited from a file' -NoNewline
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '1', '-ContentFile', './_update.md') `
            -Responses @(@{ Status = 200; Body = '{"id":1}' })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 1
        $result.Requests[0].Path | Should -Match '/threads/42/comments/1\?'
        ($result.Requests[0].Body | ConvertFrom-Json).content | Should -BeExactly 'Edited from a file'
    }

    It 'refuses to edit a comment no automated review wrote' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_update.md') -Value 'Edited'
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '2', '-ContentFile', './_update.md')

        Assert-Refused $result $run 'comment #2 in thread #42 was not written by an automated review'
    }

    It 'refuses another pull request named in the environment' {
        $run = New-PolicyRun
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-Status', 'Fixed') -Environment @{ PRID = '124' }

        Assert-Refused $result $run 'not on pull request #124'
    }

    It 'refuses a credential in the new content' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_update.md') -Value "Leaked $Secret"
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '1', '-ContentFile', './_update.md') `
            -Environment @{ CLAUDE_CODE_OAUTH_TOKEN = $Secret }

        Assert-Refused $result $run 'CLAUDE_CODE_OAUTH_TOKEN'
    }

    It 'does not resolve the thread when the -ContentFile is empty' {
        $run = New-PolicyRun
        Set-Content -LiteralPath (Join-Path $run.WorkDir '_update.md') -Value '' -NoNewline
        $result = Invoke-Update -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '1', '-ContentFile', './_update.md')

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 0
        (Get-PolicyActivity $run.StateDir).Unfinished | Should -Be 1
    }

    It 'still accepts inline -Content outside a review' {
        $run = New-PolicyRun
        $result = Invoke-PolicyScript -Script 'Update-CopilotComment.ps1' -Run $run -Standalone `
            -Arguments @('-ThreadId', '42', '-CommentId', '1', '-Content', 'Edited') -Responses @(@{ Status = 200; Body = '{"id":1}' })

        $result.ExitCode | Should -Be 0
        ($result.Requests[0].Body | ConvertFrom-Json).content | Should -BeExactly 'Edited'
    }
}

Describe 'Delete-CopilotComment.ps1 under a review policy' {

    BeforeAll {
        function Invoke-Delete {
            param($Run, [string[]]$Arguments, [hashtable[]]$Responses = @(), [hashtable]$Environment = @{}, [switch]$NoPolicy)
            $policy = if ($NoPolicy) { $null } else {
                New-TestPolicyJson -WorkspaceRoot $Run.WorkDir -Threads @(New-TestPolicyThread -ThreadId 42 -AgentCommentIds 1, 3)
            }
            Invoke-PolicyScript -Script 'Delete-CopilotComment.ps1' -Run $Run -Policy $policy -Arguments $Arguments `
                -Responses $Responses -Environment $Environment
        }
    }

    It 'deletes a comment an automated review wrote' {
        $run = New-PolicyRun
        $result = Invoke-Delete -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '3') -Responses @(@{ Status = 200; Body = '' })

        $result.ExitCode | Should -Be 0
        $result.Requests[0].Method | Should -Be 'DELETE'
        $result.Requests[0].Path | Should -Match '/threads/42/comments/3\?'
        $activity = Get-PolicyActivity $run.StateDir
        $activity.Unfinished | Should -Be 0
        $activity.Done | Should -Be @('delete')
    }

    It 'refuses to delete a comment no automated review wrote' {
        $run = New-PolicyRun
        $result = Invoke-Delete -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '2')

        Assert-Refused $result $run 'may not delete it'
    }

    It 'refuses another server named in the environment' {
        $run = New-PolicyRun
        $result = Invoke-Delete -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '1') `
            -Environment @{ AZUREDEVOPS_COLLECTION_URI = 'http://127.0.0.2:9/testorg' }

        Assert-Refused $result $run 'may act only on pull request #123'
    }

    It 'refuses every action without a policy' {
        $run = New-PolicyRun
        $result = Invoke-Delete -Run $run -Arguments @('-ThreadId', '42', '-CommentId', '1') -NoPolicy

        Assert-Refused $result $run 'missing or unreadable'
    }

    It 'treats invalid IDs as an unfinished attempt without failing' {
        $run = New-PolicyRun
        $result = Invoke-Delete -Run $run -Arguments @('-ThreadId', 'abc', '-CommentId', '1')

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 0
        (Get-PolicyActivity $run.StateDir).Unfinished | Should -Be 1
    }
}
