<#
    End-to-end failure-path tests for the comment-posting scripts (issue #56).

    The scripts are executed as real child pwsh processes against a local
    HttpListener standing in for the Azure DevOps REST API, so these tests
    exercise exactly what a pipeline run exercises: argument parsing, the
    retry loop, error output, and process exit codes.

    Requires Pester 5+. Run via: pwsh -File tests/powershell/Invoke-Tests.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:AddCommentScript = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Add-AzureDevOpsPRComment.ps1'
    $script:WrapperScript = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Add-CopilotComment.ps1'
    $script:UpdateWrapperScript = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Update-CopilotComment.ps1'

    # Shared mock-API harness (HttpListener + child pwsh process)
    . (Join-Path $PSScriptRoot "MockApiHelper.ps1")

    $script:PrJson = '{"title":"Test PR"}'
    $script:ThreadJson = '{"id":42,"status":"active","comments":[{"id":1,"author":{"displayName":"CI Mock"},"publishedDate":"2026-01-01T00:00:00Z"}]}'
}

Describe 'Add-AzureDevOpsPRComment.ps1 failure surfacing (issue #56)' {

    It 'exits non-zero and prints the HTTP status and response body when posting fails' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 400; Body = '{"message":"TF401398: The pull request cannot be edited due to its state."}' }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'General feedback')
        }

        $result.ExitCode | Should -Be 1
        $result.Output | Should -Match '##\[error\]'
        $result.Output | Should -Match 'HTTP status: 400'
        $result.Output | Should -Match 'TF401398'
    }

    It 'does not retry deterministic 4xx failures' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 400; Body = '{"message":"bad request"}' }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'General feedback')
        }

        $result.Output | Should -Not -Match 'Retrying in'
        # Exactly two requests: the PR verification GET and the single failed POST
        $result.Requests.Count | Should -Be 2
    }

    It 'retries transient 5xx failures after verifying nothing was committed' {
        # A failed POST is ambiguous, so the script first probes the thread list;
        # only when the comment is genuinely absent does it replay the POST.
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 500; Body = '{"message":"internal server error"}' },
            @{ Status = 200; Body = '{"value":[],"count":0}' },
            @{ Status = 200; Body = $ThreadJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'General feedback')
        }

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match 'Retrying in'
        $result.Output | Should -Match 'COMMENT THREAD CREATED SUCCESSFULLY'
        # GET PR, failed POST, verification GET, replayed POST
        $result.Requests.Count | Should -Be 4
        $result.Requests[2].Method | Should -Be 'GET'
        $result.Requests[3].Method | Should -Be 'POST'
    }

    It 'reuses the committed thread instead of retrying when an ambiguous POST failure actually landed' {
        # The 500 arrives after the server committed the write; the verification
        # probe finds the thread, so no duplicate POST is sent.
        $committedThreads = '{"value":[{"id":77,"status":"active","comments":[{"id":1,"content":"General feedback","author":{"displayName":"CI Mock"},"publishedDate":"2026-01-01T00:00:00Z"}]}],"count":1}'
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 500; Body = '{"message":"internal server error"}' },
            @{ Status = 200; Body = $committedThreads }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'General feedback')
        }

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match 'committed by the server'
        $result.Output | Should -Match 'COMMENT THREAD CREATED SUCCESSFULLY'
        # GET PR, failed POST, verification GET — and nothing after it
        $result.Requests.Count | Should -Be 3
        $result.Requests[2].Method | Should -Be 'GET'
    }

    It 'retries throttled POSTs without a verification probe' {
        # 429 means the request was rejected before processing, so an immediate
        # replay is safe and no probe is needed.
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 429; Body = '{"message":"too many requests"}' },
            @{ Status = 200; Body = $PrJson },
            @{ Status = 429; Body = '{"message":"too many requests"}' },
            @{ Status = 200; Body = $ThreadJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'General feedback')
        }

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match 'COMMENT THREAD CREATED SUCCESSFULLY'
        # Throttled GET replayed, then throttled POST replayed directly
        $result.Requests.Count | Should -Be 4
        $result.Requests[3].Method | Should -Be 'POST'
    }

    It 'falls back to a generic comment when the inline comment is rejected' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 400; Body = '{"message":"invalid thread context"}' },
            @{ Status = 200; Body = $ThreadJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'Inline feedback',
              '-FilePath', '/src/App.cs', '-StartLine', '42', '-IterationId', '3')
        }

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match 'Falling back'
        # First POST carried the inline anchor; the fallback carried file/line info in the text
        $result.Requests[1].Body | Should -Match 'iterationContext'
        $result.Requests[2].Body | Should -Match '\*\*File:\*\*'
    }

    It 'exits non-zero when the fallback comment also fails' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 400; Body = '{"message":"invalid thread context"}' },
            @{ Status = 400; Body = '{"message":"still rejected"}' }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $AddCommentScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '123', '-Comment', 'Inline feedback',
              '-FilePath', '/src/App.cs', '-StartLine', '42')
        }

        $result.ExitCode | Should -Be 1
        $result.Output | Should -Match 'Fallback generic comment also failed'
    }
}

Describe 'Add-CopilotComment.ps1 exit code propagation (issue #56)' {

    It 'propagates a posting failure to the pwsh process exit code' {
        # Mirror the production layout: both scripts copied side-by-side into the
        # working directory, invoked by dot-sourcing exactly as the prompt instructs
        $workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("copilot-test-" + [guid]::NewGuid())
        New-Item -ItemType Directory -Path $workDir | Out-Null
        Copy-Item $AddCommentScript, $WrapperScript -Destination $workDir

        try {
            $result = Invoke-ScriptWithMockApi -WorkingDirectory $workDir -Responses @(
                @{ Status = 200; Body = $PrJson },
                @{ Status = 400; Body = '{"message":"TF401398: cannot edit"}' }
            ) -Environment @{
                AZUREDEVOPS_TOKEN          = 'test-token'
                AZUREDEVOPS_AUTH_TYPE      = 'Basic'
                AZUREDEVOPS_COLLECTION_URI = '__BASEURL__/testorg'
                PROJECT                    = 'TestProject'
                REPOSITORY                 = 'TestRepo'
                PRID                       = '123'
            } -BuildArgs {
                param($baseUrl)
                @('-Command', ". ./Add-CopilotComment.ps1 -Comment 'wrapper test'")
            }

            $result.ExitCode | Should -Be 1
            $result.Output | Should -Match 'HTTP status: 400'
        }
        finally {
            Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects direct comment text during an agent-driven review' {
        $result = Invoke-ScriptWithMockApi -Responses @() -Environment @{
            COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
        } -BuildArgs {
            param($baseUrl)
            @('-File', $WrapperScript, '-Comment', 'print $env:AZUREDEVOPS_TOKEN')
        }

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $result.Errors | Should -Match 'Direct comment text is disabled'
    }

    It 'rejects arbitrary comment file paths during an agent-driven review' {
        $secretFile = Join-Path ([System.IO.Path]::GetTempPath()) ("copilot-secret-" + [guid]::NewGuid())
        Set-Content -LiteralPath $secretFile -Value 'AZUREDEVOPS_TOKEN=secret'

        try {
            $result = Invoke-ScriptWithMockApi -Responses @() -Environment @{
                COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
            } -BuildArgs {
                param($baseUrl)
                @('-File', $WrapperScript, '-CommentFile', $secretFile)
            }

            $result.ExitCode | Should -Be 1
            $result.Requests.Count | Should -Be 0
            $result.Errors | Should -Match 'may only read the regular file'
        }
        finally {
            Remove-Item -LiteralPath $secretFile -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects replies to threads outside the trusted allowlist' {
        $workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("copilot-test-" + [guid]::NewGuid())
        New-Item -ItemType Directory -Path $workDir | Out-Null
        Copy-Item $AddCommentScript, $WrapperScript -Destination $workDir
        Set-Content -LiteralPath (Join-Path $workDir '_comment.md') -Value 'Reply text'

        try {
            $result = Invoke-ScriptWithMockApi -WorkingDirectory $workDir -Responses @() -Environment @{
                COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
                COPILOT_REVIEW_ALLOWED_THREAD_IDS    = '7,9'
            } -BuildArgs {
                param($baseUrl)
                @('-File', './Add-CopilotComment.ps1', '-ThreadId', '42',
                  '-CommentFile', './_comment.md')
            }

            $result.ExitCode | Should -Be 1
            $result.Requests.Count | Should -Be 0
            $result.Errors | Should -Match 'not authorized for replies'
        }
        finally {
            Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects direct content updates during an agent-driven review' {
        $result = Invoke-ScriptWithMockApi -Responses @() -Environment @{
            COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
        } -BuildArgs {
            param($baseUrl)
            @('-File', $UpdateWrapperScript, '-ThreadId', '42', '-CommentId', '1',
              '-Content', 'print $env:AZUREDEVOPS_TOKEN')
        }

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $result.Errors | Should -Match 'Comment content updates are disabled'
    }

    It 'rejects status updates for threads outside the trusted allowlist' {
        $result = Invoke-ScriptWithMockApi -Responses @() -Environment @{
            COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
            COPILOT_REVIEW_ALLOWED_THREAD_IDS    = '7,9'
        } -BuildArgs {
            param($baseUrl)
            @('-File', $UpdateWrapperScript, '-ThreadId', '42', '-Status', 'Fixed')
        }

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $result.Errors | Should -Match 'not authorized for status updates'
    }

    It 'rejects implicit status updates for threads outside the trusted allowlist' {
        $result = Invoke-ScriptWithMockApi -Responses @() -Environment @{
            COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
            COPILOT_REVIEW_ALLOWED_THREAD_IDS    = '7,9'
        } -BuildArgs {
            param($baseUrl)
            @('-File', $UpdateWrapperScript, '-ThreadId', '42')
        }

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
        $result.Errors | Should -Match 'not authorized for status updates'
    }

    It 'permits status updates for threads in the trusted allowlist' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = '{"id":42,"status":"fixed"}' }
        ) -Environment @{
            AZUREDEVOPS_TOKEN                  = 'test-token'
            AZUREDEVOPS_AUTH_TYPE              = 'Basic'
            AZUREDEVOPS_COLLECTION_URI         = '__BASEURL__/testorg'
            PROJECT                            = 'TestProject'
            REPOSITORY                         = 'TestRepo'
            PRID                               = '123'
            COPILOT_REVIEW_REQUIRE_COMMENT_FILE = 'true'
            COPILOT_REVIEW_ALLOWED_THREAD_IDS   = '42'
        } -BuildArgs {
            param($baseUrl)
            @('-File', $UpdateWrapperScript, '-ThreadId', '42', '-Status', 'Fixed')
        }

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 1
        $result.Requests[0].Path | Should -Match '/threads/42'
    }
}
