BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:VoteScript = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Set-AzureDevOpsPRVote.ps1'
    . (Join-Path $PSScriptRoot 'MockApiHelper.ps1')
    $script:ReviewerId = '11111111-2222-3333-4444-555555555555'
    $script:Source = 'a' * 40
    $script:Target = 'b' * 40

    function Get-VoteResponses {
        param([int]$Previous = -10, [int]$NewVote = -5, [string]$SourceCommit = $Source, [string]$TargetCommit = $Target)
        @(
            @{ Status = 200; Body = (ConvertTo-Json -Compress @{ authenticatedUser = @{ id = $ReviewerId } }) },
            @{ Status = 200; Body = (ConvertTo-Json -Depth 4 -Compress @{
                status = 'active'; isDraft = $false; sourceRefName = 'refs/heads/feature'; targetRefName = 'refs/heads/main'
                reviewers = @(@{ id = $ReviewerId; vote = $Previous }, @{ id = '99999999-2222-3333-4444-555555555555'; vote = 10 })
            }) },
            @{ Status = 200; Body = (ConvertTo-Json -Depth 3 -Compress @{ value = @(@{ name = 'refs/heads/feature'; objectId = $SourceCommit }) }) },
            @{ Status = 200; Body = (ConvertTo-Json -Depth 3 -Compress @{ value = @(@{ name = 'refs/heads/main'; objectId = $TargetCommit }) }) },
            @{ Status = 200; Body = (ConvertTo-Json -Compress @{ id = $ReviewerId; vote = $NewVote }) }
        )
    }

    function Invoke-Vote {
        param([hashtable[]]$Responses, [int]$Vote = -5, [int]$Previous = -10, [string]$ExpectedIdentity = $ReviewerId)
        $voteScriptPath = $VoteScript
        $sourceSha = $Source
        $targetSha = $Target
        Invoke-ScriptWithMockApi -Responses $Responses -BuildArgs {
            param($baseUrl)
            @('-File', $voteScriptPath, '-Token', 'test-token', '-AuthType', 'Bearer',
              '-CollectionUri', "$baseUrl/testorg", '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '123',
              '-Vote', "$Vote", '-ExpectedVote', "$Previous", '-ExpectedReviewerId', $ExpectedIdentity,
              '-ExpectedSourceCommit', $sourceSha, '-ExpectedTargetCommit', $targetSha)
        }.GetNewClosure()
    }
}

Describe 'Pull request votes' {
    It 'updates only the review account from <Previous> to <Vote>' -ForEach @(
        @{ Previous = -10; Vote = -5 },
        @{ Previous = 10; Vote = -5 },
        @{ Previous = -5; Vote = 10 }
    ) {
        $result = Invoke-Vote -Responses (Get-VoteResponses -Previous $Previous -NewVote $Vote) -Previous $Previous -Vote $Vote
        $result.ExitCode | Should -Be 0
        $writes = @($result.Requests | Where-Object Method -eq 'PUT')
        $writes.Count | Should -Be 1
        $writes[0].Path | Should -Match "/reviewers/$ReviewerId\?api-version=7.1$"
        $body = $writes[0].Body | ConvertFrom-Json
        $body.id | Should -Be $ReviewerId
        $body.vote | Should -Be $Vote
        $result.Output | Should -Match 'Pull request vote updated:'
    }

    It 'reads the previous vote without changing it' {
        $outputFile = Join-Path $TestDrive 'snapshot.json'
        $voteScriptPath = $VoteScript
        $result = Invoke-ScriptWithMockApi -Responses (Get-VoteResponses) -BuildArgs {
            param($baseUrl)
            @('-File', $voteScriptPath, '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo', '-Id', '123', '-ReadState', '-OutputFile', $outputFile)
        }.GetNewClosure()
        $result.ExitCode | Should -Be 0
        @($result.Requests | Where-Object Method -ne 'GET').Count | Should -Be 0
        $snapshot = Get-Content -Raw $outputFile | ConvertFrom-Json
        $snapshot.reviewerId | Should -Be $ReviewerId
        $snapshot.previousVote | Should -Be -10
        $snapshot.sourceCommit | Should -Be $Source
        $snapshot.targetCommit | Should -Be $Target
    }

    It 'does not overwrite an unchanged vote' {
        $result = Invoke-Vote -Responses (Get-VoteResponses -Previous -5) -Previous -5
        $result.ExitCode | Should -Be 0
        @($result.Requests | Where-Object Method -eq 'PUT').Count | Should -Be 0
    }

    It 'reconfirms approval for the reviewed iteration when the displayed vote is unchanged' {
        $result = Invoke-Vote -Responses (Get-VoteResponses -Previous 10 -NewVote 10) -Previous 10 -Vote 10
        $result.ExitCode | Should -Be 0
        $write = @($result.Requests | Where-Object Method -eq 'PUT')
        $write.Count | Should -Be 1
        ($write[0].Body | ConvertFrom-Json).isReapprove | Should -BeTrue
    }

    It 'keeps the vote when the source or target branch changes during the review' -ForEach @('source', 'target') {
        $sourceCommit = if ($_ -eq 'source') { 'c' * 40 } else { $Source }
        $targetCommit = if ($_ -eq 'target') { 'c' * 40 } else { $Target }
        $result = Invoke-Vote -Responses (Get-VoteResponses -SourceCommit $sourceCommit -TargetCommit $targetCommit)
        $result.ExitCode | Should -Be 1
        @($result.Requests | Where-Object Method -eq 'PUT').Count | Should -Be 0
        $result.Output | Should -Match 'source or target branch changed'
    }

    It 'does not overwrite a vote changed by another run' {
        $result = Invoke-Vote -Responses (Get-VoteResponses -Previous 10)
        $result.ExitCode | Should -Be 1
        @($result.Requests | Where-Object Method -eq 'PUT').Count | Should -Be 0
    }

    It 'refuses to vote for a different reviewer' {
        $result = Invoke-Vote -Responses (Get-VoteResponses) -ExpectedIdentity '99999999-2222-3333-4444-555555555555'
        $result.ExitCode | Should -Be 1
        @($result.Requests | Where-Object Method -eq 'PUT').Count | Should -Be 0
    }
}
