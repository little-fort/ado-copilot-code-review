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

    <#
        Runs `pwsh <args>` as a child process while serving a scripted sequence
        of HTTP responses from a local listener. Returns exit code, stdout,
        stderr, and the requests the script actually made.
    #>
    function Invoke-ScriptWithMockApi {
        param(
            # Receives the mock API base URL (e.g. http://localhost:51234),
            # returns the argument array for pwsh
            [Parameter(Mandatory)] [scriptblock]$BuildArgs,
            # Ordered responses: @{ Status = <int>; Body = <string> }
            [Parameter(Mandatory)] [hashtable[]]$Responses,
            # Extra environment variables for the child process only. The token
            # __BASEURL__ in a value is replaced with the mock API base URL,
            # which is not known until the listener claims a port.
            [hashtable]$Environment = @{},
            [string]$WorkingDirectory,
            [int]$TimeoutSeconds = 90
        )

        # Find a free port, then stand the listener up on it
        $tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        $tcp.Start()
        $port = ([System.Net.IPEndPoint]$tcp.LocalEndpoint).Port
        $tcp.Stop()

        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add("http://localhost:$port/")
        $listener.Start()

        $baseUrl = "http://localhost:$port"
        $requests = [System.Collections.Generic.List[object]]::new()

        try {
            $psi = [System.Diagnostics.ProcessStartInfo]::new()
            $psi.FileName = 'pwsh'
            $psi.ArgumentList.Add('-NoProfile')
            foreach ($arg in (& $BuildArgs $baseUrl)) { $psi.ArgumentList.Add([string]$arg) }
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            if ($WorkingDirectory) { $psi.WorkingDirectory = $WorkingDirectory }
            foreach ($key in $Environment.Keys) {
                $psi.Environment[$key] = ([string]$Environment[$key]) -replace '__BASEURL__', $baseUrl
            }

            $proc = [System.Diagnostics.Process]::Start($psi)
            # Async pumps prevent the child blocking on a full stdout/stderr pipe
            $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
            $stderrTask = $proc.StandardError.ReadToEndAsync()

            $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
            $responseIndex = 0
            $ctxTask = $listener.GetContextAsync()

            while (-not $proc.HasExited -and [DateTime]::UtcNow -lt $deadline) {
                if ($ctxTask.Wait(200)) {
                    $ctx = $ctxTask.Result
                    $reader = [System.IO.StreamReader]::new($ctx.Request.InputStream)
                    $requests.Add([pscustomobject]@{
                        Method = $ctx.Request.HttpMethod
                        Path   = $ctx.Request.Url.PathAndQuery
                        Body   = $reader.ReadToEnd()
                    })

                    $resp = if ($responseIndex -lt $Responses.Count) {
                        $Responses[$responseIndex]
                    } else {
                        @{ Status = 500; Body = '{"message":"mock API: unexpected extra request"}' }
                    }
                    $responseIndex++

                    $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$resp.Body)
                    $ctx.Response.StatusCode = [int]$resp.Status
                    $ctx.Response.ContentType = 'application/json'
                    $ctx.Response.ContentLength64 = $bytes.Length
                    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                    $ctx.Response.Close()

                    $ctxTask = $listener.GetContextAsync()
                }
            }

            if (-not $proc.HasExited) {
                $proc.Kill()
                throw "Script under test did not exit within $TimeoutSeconds seconds."
            }
            $proc.WaitForExit()

            return [pscustomobject]@{
                ExitCode = $proc.ExitCode
                Output   = $stdoutTask.Result
                Errors   = $stderrTask.Result
                Requests = $requests
            }
        }
        finally {
            $listener.Stop()
            $listener.Close()
        }
    }

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

    It 'retries transient 5xx failures and succeeds' {
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 500; Body = '{"message":"internal server error"}' },
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
        $result.Requests.Count | Should -Be 3
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
}
