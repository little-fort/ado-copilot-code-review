<#
    Tests for Get-AzureDevOpsPRChanges.ps1 pagination (issue #64) and log
    sanitization (issue #69).

    The iteration-changes endpoint returns 100 entries per page by default
    and reports the next page via nextSkip/nextTop; the PR-commits endpoint
    pages via an x-ms-continuationtoken response header. Before #64 the
    script made a single request to each, silently capping large PRs. The
    script runs as a real child pwsh process against a local HttpListener
    standing in for the Azure DevOps REST API.

    Requires Pester 5+. Run via: pwsh -File tests/powershell/Invoke-Tests.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:ChangesScript = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Get-AzureDevOpsPRChanges.ps1'

    # Shared mock-API harness (HttpListener + child pwsh process)
    . (Join-Path $PSScriptRoot 'MockApiHelper.ps1')

    $script:PrJson = '{"pullRequestId":5,"title":"Big refactor","status":"active"}'
    $script:IterationsJson = '{"count":1,"value":[{"id":3,"createdDate":"2026-01-01T00:00:00Z","updatedDate":"2026-01-02T00:00:00Z","sourceRefCommit":{"commitId":"1111111111111111111111111111111111111111"},"targetRefCommit":{"commitId":"2222222222222222222222222222222222222222"}}]}'

    # Builds a commits response page; each number becomes a 40-digit commit ID
    function New-CommitsPage {
        param([int[]]$Numbers, [string]$MessagePrefix = 'Commit')
        @{
            count = $Numbers.Count
            value = @($Numbers | ForEach-Object {
                @{
                    commitId = ('{0:D40}' -f $_)
                    comment  = "$MessagePrefix $_"
                    author   = @{ name = 'Dev'; date = '2026-01-01T00:00:00Z' }
                }
            })
        } | ConvertTo-Json -Depth 5
    }

    # Builds an iteration-changes response page. Pass -OmitPaging to mimic
    # servers whose responses carry no nextSkip/nextTop fields.
    function New-ChangesPage {
        param([string[]]$Paths, [int]$NextSkip = 0, [int]$NextTop = 0, [switch]$OmitPaging)
        $page = @{
            changeEntries = @($Paths | ForEach-Object { @{ changeType = 'edit'; item = @{ path = $_ } } })
        }
        if (-not $OmitPaging) {
            $page.nextSkip = $NextSkip
            $page.nextTop = $NextTop
        }
        $page | ConvertTo-Json -Depth 5
    }

    # Runs the script against the given responses; adds the output file text
    function Invoke-ChangesScript {
        param([hashtable[]]$Responses)

        $outDir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $outDir | Out-Null
        $outFile = Join-Path $outDir 'Iteration_Details.txt'
        $scriptPath = $ChangesScript

        $result = Invoke-ScriptWithMockApi -Responses $Responses -BuildArgs {
            param($baseUrl)
            @('-File', $scriptPath,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '5', '-OutputFile', $outFile)
        }.GetNewClosure()

        $fileText = if (Test-Path $outFile) { Get-Content -Path $outFile -Raw } else { '' }
        $result | Add-Member -NotePropertyName FileText -NotePropertyValue $fileText -PassThru
    }
}

Describe 'Get-AzureDevOpsPRChanges.ps1 pagination (issue #64)' {

    It 'follows the commits continuation token and the changes nextSkip/nextTop values across pages' {
        $result = Invoke-ChangesScript -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $IterationsJson },
            @{ Status = 200; Body = (New-CommitsPage 1, 2); Headers = @{ 'x-ms-continuationtoken' = 'token/2+=' } },
            @{ Status = 200; Body = (New-CommitsPage 3) },
            @{ Status = 200; Body = (New-ChangesPage '/src/a.cs', '/src/b.cs' -NextSkip 2 -NextTop 2) },
            @{ Status = 200; Body = (New-ChangesPage '/src/c.cs') }
        )

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 6

        # The token goes back URL-encoded on the second commits request
        $result.Requests[2].Path | Should -Not -Match 'continuationToken='
        $result.Requests[3].Path | Should -Match ([regex]::Escape('continuationToken=token%2F2%2B%3D'))

        # The first changes request is unchanged from the pre-pagination
        # request; the second uses the server-provided paging values
        $result.Requests[4].Path | Should -Not -Match 'skip='
        $result.Requests[5].Path | Should -Match '(\$|%24)top=2&(\$|%24)skip=2'

        $result.FileText | Should -Match 'Total commits: 3'
        $result.FileText | Should -Match 'Total files changed: 3'
        foreach ($path in '/src/a.cs', '/src/b.cs', '/src/c.cs') {
            $result.FileText | Should -Match ([regex]::Escape("[Modified] $path"))
        }
        $result.FileText | Should -Not -Match 'INCOMPLETE'
    }

    It 'makes one request per endpoint when responses carry no paging information' {
        # Older servers omit nextSkip/nextTop; no continuation header means
        # the commit list is complete
        $result = Invoke-ChangesScript -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $IterationsJson },
            @{ Status = 200; Body = (New-CommitsPage 1) },
            @{ Status = 200; Body = (New-ChangesPage '/src/a.cs' -OmitPaging) }
        )

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 4
        $result.FileText | Should -Match 'Total commits: 1'
        $result.FileText | Should -Match 'Total files changed: 1'
    }

    It 'marks the file list incomplete when a later page fails' {
        $result = Invoke-ChangesScript -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $IterationsJson },
            @{ Status = 200; Body = (New-CommitsPage 1) },
            @{ Status = 200; Body = (New-ChangesPage '/src/a.cs', '/src/b.cs' -NextSkip 2 -NextTop 2) },
            @{ Status = 500; Body = '{"message":"internal server error"}' }
        )

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match '##\[warning\]Only 2 changed file'
        $result.FileText | Should -Match 'Files retrieved: 2 \(INCOMPLETE'
        $result.FileText | Should -Not -Match 'Total files changed'
    }

    It 'reports a failed file-list request instead of claiming there are no changes' {
        $result = Invoke-ChangesScript -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $IterationsJson },
            @{ Status = 200; Body = (New-CommitsPage 1) },
            @{ Status = 500; Body = '{"message":"internal server error"}' }
        )

        $result.FileText | Should -Match 'Changed files could not be retrieved'
        $result.FileText | Should -Not -Match 'No file changes found'
    }

    It 'stops when the server repeats a continuation token' {
        $result = Invoke-ChangesScript -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = $IterationsJson },
            @{ Status = 200; Body = (New-CommitsPage 1); Headers = @{ 'x-ms-continuationtoken' = 'same' } },
            @{ Status = 200; Body = (New-CommitsPage 2); Headers = @{ 'x-ms-continuationtoken' = 'same' } },
            @{ Status = 200; Body = (New-ChangesPage '/src/a.cs') }
        )

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 5
        $result.FileText | Should -Match 'Commits retrieved: 2 \(INCOMPLETE'
    }
}

Describe 'Get-AzureDevOpsPRChanges.ps1 log output (issue #69)' {

    It 'neutralizes logging commands in the PR title and commit messages on the console only' {
        $prJson = '{"pullRequestId":5,"title":"Title ##vso[task.setvariable variable=Leak]x","status":"active"}'
        $result = Invoke-ChangesScript -Responses @(
            @{ Status = 200; Body = $prJson },
            @{ Status = 200; Body = $IterationsJson },
            @{ Status = 200; Body = (New-CommitsPage 1 -MessagePrefix 'fix ##vso[task.complete result=Failed;]') },
            @{ Status = 200; Body = (New-ChangesPage '/src/a.cs') }
        )

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Not -Match '##vso\['
        $result.Output | Should -Match ([regex]::Escape('# #vso[task.complete result=Failed;]'))
        $result.FileText | Should -Match ([regex]::Escape('fix ##vso[task.complete result=Failed;]'))
    }
}
