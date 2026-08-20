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
        # Ordered responses: @{ Status = <int>; Body = <string> }. An empty
        # array is valid for tests asserting the script makes no API calls.
        [Parameter(Mandatory)] [AllowEmptyCollection()] [hashtable[]]$Responses,
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
