<#
    Shared test harness: runs a script under test as a real child pwsh process
    while serving a scripted sequence of HTTP responses from a local
    HttpListener. Dot-source this file from a Pester BeforeAll block.

    Used by the comment-posting tests (mocking the Azure DevOps REST API) and
    the Jira work-item tests (mocking the Jira Cloud REST API).
#>

function Invoke-ScriptWithMockApi {
    param(
        # Receives the mock API base URL (e.g. http://localhost:51234),
        # returns the argument array for pwsh
        [Parameter(Mandatory)] [scriptblock]$BuildArgs,
        # Ordered responses: @{ Status = <int>; Body = <string> }, plus an
        # optional Headers = @{ Name = 'value' } for response headers (e.g.
        # continuation tokens). An empty array is valid for tests asserting
        # the script makes no API calls.
        [Parameter(Mandatory)] [AllowEmptyCollection()] [hashtable[]]$Responses,
        # Extra environment variables for the child process only. The token
        # __BASEURL__ in a value is replaced with the mock API base URL,
        # which is not known until the listener claims a port.
        [hashtable]$Environment = @{},
        [string]$WorkingDirectory,
        [int]$TimeoutSeconds = 90,
        # Files to write before the script starts: path => content, with
        # __BASEURL__ replaced as above (e.g. a review policy whose
        # destination is the mock API)
        [hashtable]$Files = @{}
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
        foreach ($filePath in $Files.Keys) {
            [System.IO.File]::WriteAllText($filePath, ([string]$Files[$filePath]).Replace('__BASEURL__', $baseUrl),
                [System.Text.UTF8Encoding]::new($false))
        }

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
                    Method      = $ctx.Request.HttpMethod
                    Path        = $ctx.Request.Url.PathAndQuery
                    ContentType = $ctx.Request.ContentType
                    Body        = $reader.ReadToEnd()
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
                if ($resp.Headers) {
                    foreach ($headerName in $resp.Headers.Keys) {
                        $ctx.Response.AddHeader($headerName, [string]$resp.Headers[$headerName])
                    }
                }
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

<#
    Review policy fixtures (PR #65 rework). The task writes review-policy.json
    to its state folder before the agent runs, and the comment scripts refuse
    every action when the folder is set but the policy is missing. Pass the
    JSON to Invoke-ScriptWithMockApi -Files so the destination can name the
    mock API (__BASEURL__) before the script starts.
#>
function New-TestPolicyThread {
    param(
        [int]$ThreadId,
        [bool]$MayAct = $true,
        [int[]]$AgentCommentIds = @(),
        [bool]$AgentAuthored = $true,
        [string]$Status = 'active'
    )

    [ordered]@{
        threadId        = $ThreadId
        status          = $Status
        agentAuthored   = $AgentAuthored
        mayReply        = $MayAct
        mayUpdateStatus = $MayAct
        agentCommentIds = @($AgentCommentIds)
    }
}

function New-TestPolicyJson {
    param(
        [Parameter(Mandatory)] [string]$WorkspaceRoot,
        [object[]]$Threads = @(),
        [bool]$ThreadsComplete = $true,
        [string[]]$ChangedFiles = @('/src/App.cs'),
        [bool]$ChangedFilesComplete = $true,
        [string]$CollectionUri = '__BASEURL__/testorg',
        [string]$Project = 'TestProject',
        [string]$Repository = 'TestRepo',
        [int]$PullRequestId = 123
    )

    ConvertTo-Json -Depth 6 -InputObject ([ordered]@{
        schemaVersion        = 1
        destination          = [ordered]@{
            collectionUri = $CollectionUri
            project       = $Project
            repository    = $Repository
            pullRequestId = $PullRequestId
        }
        resolveThreads       = 'all'
        rawPrompt            = $false
        workspaceRoot        = $WorkspaceRoot
        threadsComplete      = $ThreadsComplete
        threads              = @($Threads)
        changedFilesComplete = $ChangedFilesComplete
        changedFiles         = @($ChangedFiles)
    })
}
