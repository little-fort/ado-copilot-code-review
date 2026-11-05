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
    # Runs the script against the given responses; adds the output file text.
    # With -StateKeyPrefix the review-state parameters are passed too, and the
    # written Review_State.json is returned (raw text and parsed).
    function Invoke-ChangesScript {
        param([hashtable[]]$Responses, [string]$StateKeyPrefix)

        $outDir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $outDir | Out-Null
        $outFile = Join-Path $outDir 'Iteration_Details.txt'
        $stateFile = Join-Path $outDir 'state/Review_State.json'
        $scriptPath = $ChangesScript
        $stateArgs = if ($StateKeyPrefix) { @('-ReviewStateKeyPrefix', $StateKeyPrefix, '-ReviewStateFile', $stateFile) } else { @() }

        $result = Invoke-ScriptWithMockApi -Responses $Responses -BuildArgs {
            param($baseUrl)
            @('-File', $scriptPath,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '5', '-OutputFile', $outFile) + $stateArgs
        }.GetNewClosure()

        $fileText = if (Test-Path $outFile) { Get-Content -Path $outFile -Raw } else { '' }
        $stateText = if (Test-Path $stateFile) { Get-Content -Path $stateFile -Raw } else { $null }
        $result | Add-Member -NotePropertyName FileText -NotePropertyValue $fileText
        $result | Add-Member -NotePropertyName StateText -NotePropertyValue $stateText
        $result | Add-Member -NotePropertyName State -NotePropertyValue $(if ($stateText) { $stateText | ConvertFrom-Json } else { $null }) -PassThru
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

Describe 'Get-AzureDevOpsPRChanges.ps1 review state (issue #67)' {

    BeforeAll {
        $script:Prefix = 'CopilotCodeReview.v1.0123456789abcdef'
        $script:ShaA = 'a' * 40
        $script:ShaB = 'b' * 40
        $script:ShaC = 'c' * 40
        $script:BaseOne = 'e' * 40
        $script:BaseTwo = 'f' * 40

        # Three iterations; the current one (3) can change merge base or reason
        function New-IterationsJson {
            param([string]$CurrentBase = $BaseOne, [string]$CurrentReason = 'push', [string]$SecondCreated = '2026-11-04T23:15:33.123Z')
            @{
                count = 3
                value = @(
                    @{ id = 1; reason = 'create'; sourceRefCommit = @{ commitId = $ShaA }; targetRefCommit = @{ commitId = '1' * 40 }; commonRefCommit = @{ commitId = $BaseOne } },
                    @{ id = 2; reason = 'push'; createdDate = $SecondCreated; sourceRefCommit = @{ commitId = $ShaB }; targetRefCommit = @{ commitId = '1' * 40 }; commonRefCommit = @{ commitId = $BaseOne } },
                    @{ id = 3; reason = $CurrentReason; sourceRefCommit = @{ commitId = $ShaC }; targetRefCommit = @{ commitId = '1' * 40 }; commonRefCommit = @{ commitId = $CurrentBase } }
                )
            } | ConvertTo-Json -Depth 5
        }

        # Pull request properties holding review records (iteration -> SHA),
        # next to a system property every PR has
        function New-PropertiesJson {
            param([hashtable]$Records = @{})
            $value = [ordered]@{ 'Microsoft.Git.PullRequest.SourceRefName' = @{ '$type' = 'System.String'; '$value' = 'refs/heads/feature' } }
            foreach ($iteration in $Records.Keys) {
                $value["$Prefix.Iteration.$iteration"] = @{ '$type' = 'System.String'; '$value' = $Records[$iteration] }
            }
            @{ count = $value.Count; value = $value } | ConvertTo-Json -Depth 5
        }

        # A changes page from @{ Path; Type; From } entries
        function New-DeltaPage {
            param([object[]]$Entries, [int]$NextSkip = 0, [int]$NextTop = 0)
            @{
                changeEntries = @($Entries | ForEach-Object {
                    $entry = @{ changeType = $_.Type; item = @{ path = $_.Path } }
                    if ($_.From) { $entry.originalPath = $_.From }
                    $entry
                })
                nextSkip      = $NextSkip
                nextTop       = $NextTop
            } | ConvertTo-Json -Depth 5
        }

        # The pull request's current files: an edit and a rename
        $script:FullChanges = New-DeltaPage @(
            @{ Path = '/src/a.cs'; Type = 'edit' },
            @{ Path = '/src/new.cs'; Type = 'rename'; From = '/src/old.cs' }
        )

        # Responses up to and including the properties read
        function Get-BaseResponses {
            param([string]$Iterations = (New-IterationsJson), [hashtable]$Properties)
            @(
                @{ Status = 200; Body = $PrJson },
                @{ Status = 200; Body = $Iterations },
                @{ Status = 200; Body = (New-CommitsPage 1) },
                @{ Status = 200; Body = $FullChanges },
                $Properties
            )
        }
    }

    It 'reports no state when nothing was recorded' {
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson) })

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 5
        $result.State.stateReadStatus | Should -Be 'NotFound'
        $result.State.currentIteration | Should -Be 3
        $result.State.currentSourceCommit | Should -Be $ShaC
        $result.State.fullChangesStatus | Should -Be 'Complete'
        $result.FileText | Should -Not -Match 'Changes Since Last Review'
        # Every current path, for the review policy's inline-comment check
        # (PR #65 rework); rename origins are not anchor targets
        @($result.State.changedFiles) | Should -Be @('/src/a.cs', '/src/new.cs')
    }

    It 'treats a missing properties resource as no state' {
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses (Get-BaseResponses -Properties @{ Status = 404; Body = '{"message":"not found"}' })

        $result.ExitCode | Should -Be 0
        $result.State.stateReadStatus | Should -Be 'NotFound'
        $result.Errors | Should -BeNullOrEmpty
    }

    It 'recognizes the current iteration as already reviewed without fetching a delta' {
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 3 = $ShaC }) })

        $result.Requests.Count | Should -Be 5
        $result.State.stateReadStatus | Should -Be 'Found'
        $result.State.lastReviewedIteration | Should -Be 3
        $result.State.lastReviewedSourceCommit | Should -Be $ShaC
    }

    It 'lists the changes since the last reviewed iteration' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 1 = $ShaA; 2 = $ShaB }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(
                @{ Path = '/src/a.cs'; Type = 'edit' },
                @{ Path = '/src/reverted.cs'; Type = 'edit' },
                @{ Path = '/src/new.cs'; Type = 'rename'; From = '/src/old.cs' }
            )) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.ExitCode | Should -Be 0
        $result.Requests.Count | Should -Be 6
        $result.Requests[5].Path | Should -Match '/iterations/3/changes\?.*(\$|%24)compareTo=2'
        $result.State.lastReviewedIteration | Should -Be 2
        $result.State.delta.status | Should -Be 'Complete'
        # The file the pull request reverted is dropped; the rename is kept
        @($result.State.delta.files | ForEach-Object { $_.path }) | Should -Be @('/src/a.cs', '/src/new.cs')
        ($result.State.delta.files | Where-Object { $_.path -eq '/src/new.cs' }).originalPath | Should -Be '/src/old.cs'

        $result.FileText | Should -Match '\[Changes Since Last Review\]'
        $result.FileText | Should -Match 'Last reviewed iteration: #2'
        $result.FileText | Should -Match ([regex]::Escape("git diff $ShaB $ShaC"))
        $result.FileText | Should -Match ([regex]::Escape('(from: /src/old.cs)'))
        $result.FileText | Should -Not -Match 'reverted\.cs'
    }

    It 'ignores a record whose commit does not match its iteration' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 1 = $ShaA; 2 = ('9' * 40) }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.State.lastReviewedIteration | Should -Be 1
        $result.Requests[5].Path | Should -Match '(\$|%24)compareTo=1'
    }

    It 'names the build that reviewed the last iteration so the run can be found' {
        $record = '{"sourceCommit":"' + $ShaB + '","buildId":366,"buildNumber":"20261104.1","buildProject":"My Project","recordedAt":"2026-11-04T23:21:09Z"}'
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 1 = $ShaA; 2 = $record }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.ExitCode | Should -Be 0
        $result.State.lastReviewedIteration | Should -Be 2
        $result.Output | Should -Match 'Last reviewed iteration: #2'
        # UTC regardless of the machine's time zone (Invoke-RestMethod hands back local [datetime] values)
        $result.Output | Should -Match ([regex]::Escape('Source commit:  bbbbbbbb (iteration created 2026-11-04 23:15 UTC)'))
        $result.Output | Should -Match ([regex]::Escape('Reviewed by:    build 20261104.1, 2026-11-04 23:21 UTC'))
        $result.Output | Should -Match '/testorg/My%20Project/_build/results\?buildId=366'
        $result.Output | Should -Match ([regex]::Escape("Record:         $Prefix.Iteration.2"))

        $result.State.lastReviewedRecord.name | Should -Be "$Prefix.Iteration.2"
        $result.State.lastReviewedRecord.buildId | Should -Be 366
        $result.State.lastReviewedRecord.buildNumber | Should -Be '20261104.1'
        $result.State.lastReviewedRecord.buildUrl | Should -Match '^http://.+/testorg/My%20Project/_build/results\?buildId=366$'
        # Read the timestamps from the raw text: ConvertFrom-Json would turn them into [datetime]
        $result.StateText | Should -Match '"recordedAt":\s*"2026-11-04T23:21:09Z"'
        $result.StateText | Should -Match '"iterationCreated":\s*"2026-11-04T23:15:33Z"'
    }

    It 'reports the iteration time in UTC when the server sends an offset' {
        # Invoke-RestMethod turns an offset timestamp into a local-time [datetime];
        # 'Z' timestamps (what Azure DevOps sends today) stay UTC either way
        $responses = (Get-BaseResponses -Iterations (New-IterationsJson -SecondCreated '2026-11-04T17:15:33.123-06:00') `
            -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.Output | Should -Match ([regex]::Escape('(iteration created 2026-11-04 23:15 UTC)'))
        $result.StateText | Should -Match '"iterationCreated":\s*"2026-11-04T23:15:33Z"'
    }

    It 'says when a record predates the build details' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.State.lastReviewedIteration | Should -Be 2
        $result.Output | Should -Match 'Reviewed by:    not recorded \(this record predates build details\)'
        $result.State.lastReviewedRecord.buildId | Should -BeNullOrEmpty
        $result.State.lastReviewedRecord.buildUrl | Should -BeNullOrEmpty
    }

    It 'ignores a <Name> record and falls back to the earlier one' -ForEach @(
        @{ Name = 'malformed'; Value = '{"sourceCommit":' },
        @{ Name = 'JSON array'; Value = '["' + ('b' * 40) + '"]' },
        @{ Name = 'mismatched JSON'; Value = '{"sourceCommit":"' + ('9' * 40) + '","buildId":366}' },
        @{ Name = 'commit-less'; Value = '{"buildId":366,"buildNumber":"20261104.1"}' }
    ) {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 1 = $ShaA; 2 = $Value }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.ExitCode | Should -Be 0
        $result.State.lastReviewedIteration | Should -Be 1
        $result.Requests[5].Path | Should -Match '(\$|%24)compareTo=1'
    }

    It 'shows only checked build details from a forged record' {
        # Anyone who can contribute to the pull request can write its properties
        $record = (@{
            sourceCommit = $ShaB
            buildId      = 7
            buildNumber  = '##vso[task.setvariable variable=hijack]1'
            buildProject = '../../evil.example.com'
        } | ConvertTo-Json -Compress)
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $record }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.State.lastReviewedIteration | Should -Be 2
        $result.Output | Should -Not -Match '##vso\['
        $result.Output | Should -Match ([regex]::Escape('Reviewed by:    build # #vso[task.setvariable variable=hijack]1'))
        # The link stays on the configured server: the project is escaped into one path segment
        $result.State.lastReviewedRecord.buildUrl | Should -Match '^http://localhost:\d+/testorg/\.\.%2F\.\.%2Fevil\.example\.com/_build/results\?buildId=7$'
    }

    It 'drops a build detail with the wrong shape (<Name>)' -ForEach @(
        @{ Name = 'line break in the build number'; Field = 'buildNumber'; Json = '"buildId":7,"buildNumber":"20261104.1\nReviewed by: build 1"' },
        @{ Name = 'build ID as text'; Field = 'buildId'; Json = '"buildId":"7","buildNumber":"20261104.1"' },
        @{ Name = 'negative build ID'; Field = 'buildId'; Json = '"buildId":-7' },
        @{ Name = 'unparseable time'; Field = 'recordedAt'; Json = '"recordedAt":"yesterday"' }
    ) {
        $record = '{"sourceCommit":"' + $ShaB + '",' + $Json + '}'
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $record }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        # The record still counts (its commit matches); only the bad detail is dropped
        $result.State.lastReviewedIteration | Should -Be 2
        $result.State.lastReviewedRecord.$Field | Should -BeNullOrEmpty
        $result.Output | Should -Not -Match 'Reviewed by: build 1'
    }

    It 'omits the delta when the merge base changed (rebase or target merge)' {
        $responses = Get-BaseResponses -Iterations (New-IterationsJson -CurrentBase $BaseTwo) -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.Requests.Count | Should -Be 5
        $result.State.lastReviewedIteration | Should -Be 2
        $result.State.delta.status | Should -Be 'NotApplicable'
        $result.State.delta.reason | Should -Match 'merge base'
        $result.FileText | Should -Not -Match 'Changes Since Last Review'
    }

    It 'omits the delta after a retarget' {
        $responses = Get-BaseResponses -Iterations (New-IterationsJson -CurrentReason 'retarget') -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.Requests.Count | Should -Be 5
        $result.State.delta.reason | Should -Match 'retargeted'
    }

    It 'omits the delta when the full change list is incomplete' {
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses @(
            @{ Status = 200; Body = $PrJson },
            @{ Status = 200; Body = (New-IterationsJson) },
            @{ Status = 200; Body = (New-CommitsPage 1) },
            @{ Status = 200; Body = (New-ChangesPage '/src/a.cs' -NextSkip 1 -NextTop 1) },
            @{ Status = 500; Body = '{"message":"internal server error"}' },
            @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }
        )

        $result.Requests.Count | Should -Be 6
        $result.State.fullChangesStatus | Should -Be 'Partial'
        # What was retrieved, kept as a JSON array even with a single entry
        $result.StateText | Should -Match '"changedFiles":\s*\[\s*"/src/a\.cs"\s*\]'
        $result.State.delta.status | Should -Be 'NotApplicable'
        $result.State.delta.reason | Should -Match 'incomplete'
    }

    It 'reports a failed delta fetch without failing the script' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }) + @(
            @{ Status = 500; Body = '{"message":"internal server error"}' }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.ExitCode | Should -Be 0
        $result.State.delta.status | Should -Be 'Failed'
        $result.FileText | Should -Not -Match 'Changes Since Last Review'
    }

    It 'follows paging in the changes since the last review' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' }) -NextSkip 1 -NextTop 1) },
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/new.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.Requests.Count | Should -Be 7
        $result.Requests[6].Path | Should -Match '(\$|%24)compareTo=2'
        $result.Requests[6].Path | Should -Match '(\$|%24)skip=1'
        @($result.State.delta.files).Count | Should -Be 2
    }

    It 'warns and falls back to a full review when the review state cannot be read' {
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses (Get-BaseResponses -Properties @{ Status = 403; Body = '{"message":"forbidden"}' })

        $result.ExitCode | Should -Be 0
        $result.State.stateReadStatus | Should -Be 'Failed'
        $result.Output | Should -Match '##\[warning\]Could not read the incremental review state'
        # An optional read must not paint red errors in the log
        $result.Errors | Should -BeNullOrEmpty
    }

    It 'keeps a single changed file a JSON array in the state file' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/a.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.StateText | Should -Match '"files":\s*\['
    }

    It 'says so when no pull request file changed since the review' {
        $responses = (Get-BaseResponses -Properties @{ Status = 200; Body = (New-PropertiesJson @{ 2 = $ShaB }) }) + @(
            @{ Status = 200; Body = (New-DeltaPage @(@{ Path = '/src/reverted.cs'; Type = 'edit' })) }
        )
        $result = Invoke-ChangesScript -StateKeyPrefix $Prefix -Responses $responses

        $result.State.delta.status | Should -Be 'Complete'
        @($result.State.delta.files).Count | Should -Be 0
        $result.FileText | Should -Match 'No files in this pull request changed since that review'
    }

    It 'requires both review-state parameters together' {
        $scriptPath = $ChangesScript
        $result = Invoke-ScriptWithMockApi -Responses @() -BuildArgs {
            param($baseUrl)
            @('-File', $scriptPath, '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '5',
              '-ReviewStateKeyPrefix', 'CopilotCodeReview.v1.0123456789abcdef')
        }.GetNewClosure()

        $result.ExitCode | Should -Be 1
        $result.Requests.Count | Should -Be 0
    }
}
