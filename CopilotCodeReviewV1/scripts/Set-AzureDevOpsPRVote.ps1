<#
.SYNOPSIS
    Sets the review account's vote on a pull request in Azure DevOps.

.DESCRIPTION
    This script uses the Azure DevOps REST API to read or update a pull request
    vote. Supports both PAT and OAuth/System.AccessToken authentication.

    With ReadState, saves the current vote and source/target commits to a JSON
    file. When setting a vote, verifies that the branches and previous vote have
    not changed since the review started. Only updates the vote of the account
    used for authentication; other reviewers' votes are never changed.

.PARAMETER Token
    Required. Authentication token for Azure DevOps. Can be a PAT or OAuth token.

.PARAMETER AuthType
    Optional. 'Basic' for PAT, 'Bearer' for OAuth/System.AccessToken. Default is 'Basic'.

.PARAMETER CollectionUri
    Required. The Azure DevOps collection URI.

.PARAMETER Project
    Required. The Azure DevOps project name or ID.

.PARAMETER Repository
    Required. The repository name or ID where the pull request exists.

.PARAMETER Id
    Required. The pull request ID.

.PARAMETER ReadState
    Reads the review account's current vote and PR commits without changing the vote.

.PARAMETER OutputFile
    Required with ReadState. Path to the output JSON file.

.PARAMETER Vote
    Required when setting a vote. Valid values: 10 (Approved), 5 (Approved with
    suggestions), -5 (Wait for author), -10 (Rejected).

.PARAMETER ExpectedReviewerId
    Required when setting a vote. ID of the review account at the start of the review.

.PARAMETER ExpectedVote
    Required when setting a vote. The vote at the start of the review.

.PARAMETER ExpectedSourceCommit
    Required when setting a vote. Source commit SHA reviewed by the agent.

.PARAMETER ExpectedTargetCommit
    Required when setting a vote. Target commit SHA reviewed by the agent.

.EXAMPLE
    .\Set-AzureDevOpsPRVote.ps1 -Token $token -AuthType Bearer -CollectionUri $uri -Project $project -Repository $repo -Id 123 -ReadState -OutputFile ./Vote_Snapshot.json
    Saves the current vote and PR commits before starting the review.
#>

[CmdletBinding(DefaultParameterSetName = 'Read')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    [ValidateSet('Basic', 'Bearer')]
    [string]$AuthType = 'Basic',

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CollectionUri,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Project,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Id,

    [Parameter(Mandatory = $true, ParameterSetName = 'Read')]
    [switch]$ReadState,

    [Parameter(Mandatory = $true, ParameterSetName = 'Read')]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFile,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set')]
    [ValidateSet(-10, -5, 5, 10)]
    [int]$Vote,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set')]
    [guid]$ExpectedReviewerId,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set')]
    [ValidateSet(-10, -5, 0, 5, 10)]
    [int]$ExpectedVote,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set')]
    [ValidatePattern('^[0-9a-fA-F]{40}$')]
    [string]$ExpectedSourceCommit,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set')]
    [ValidatePattern('^[0-9a-fA-F]{40}$')]
    [string]$ExpectedTargetCommit
)

#region Helper Functions

function Get-VoteLabel {
    param([int]$Value)
    switch ($Value) {
        10 { 'Approved' }
        5 { 'Approved with suggestions' }
        0 { 'No vote' }
        -5 { 'Wait for author' }
        -10 { 'Rejected' }
    }
}

function Get-AuthorizationHeader {
    param([string]$Token, [string]$AuthType = 'Basic')
    $authorization = if ($AuthType -eq 'Bearer') {
        "Bearer $Token"
    } else {
        'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Token"))
    }
    return @{ Authorization = $authorization; 'Content-Type' = 'application/json' }
}

function Invoke-AzureDevOpsApi {
    param([string]$Uri, [hashtable]$Headers, [string]$Method = 'Get', [string]$Body)
    try {
        $arguments = @{ Uri = $Uri; Headers = $Headers; Method = $Method; ErrorAction = 'Stop' }
        if ($Body) { $arguments.Body = $Body }
        return Invoke-RestMethod @arguments
    } catch {
        Write-Host "##[error]Azure DevOps PR voting API call failed ($Method)."
        if ($_.Exception.Response) {
            $status = [int]$_.Exception.Response.StatusCode
            Write-Host "##[error]HTTP status: $status"
            if ($status -eq 401 -or $status -eq 403) {
                Write-Host "##[error]PR voting requires Code (Read & write) and 'Contribute to pull requests' for the authenticated identity."
            }
        }
        # Response bodies may contain credentials or untrusted logging commands.
        throw 'Could not access the Azure DevOps PR voting API.'
    }
}

function Get-BranchCommit {
    param([string]$RefName)
    if (-not $RefName.StartsWith('refs/heads/')) {
        throw 'The PR branch reference is invalid.'
    }
    $filter = [Uri]::EscapeDataString($RefName.Substring(5))
    $refs = Invoke-AzureDevOpsApi -Uri "$repositoryUrl/refs?filter=$filter&api-version=7.1" -Headers $headers
    $branchRefs = @($refs.value | Where-Object { $_.name -ceq $RefName })
    if ($branchRefs.Count -ne 1 -or [string]$branchRefs[0].objectId -notmatch '^[0-9a-fA-F]{40}$') {
        throw 'Could not determine the live PR branch commit.'
    }
    return ([string]$branchRefs[0].objectId).ToLowerInvariant()
}

#endregion

#region Main Logic

$ErrorActionPreference = 'Stop'
$collection = $CollectionUri.TrimEnd('/')
$repositoryUrl = "$collection/$([Uri]::EscapeDataString($Project))/_apis/git/repositories/$([Uri]::EscapeDataString($Repository))"
$prUrl = "$repositoryUrl/pullRequests/$Id"
$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType

try {
    $connection = Invoke-AzureDevOpsApi -Uri "$collection/_apis/connectionData?api-version=7.1-preview.1" -Headers $headers
    $reviewerId = [guid]::Empty
    if (-not [guid]::TryParse([string]$connection.authenticatedUser.id, [ref]$reviewerId) -or $reviewerId -eq [guid]::Empty) {
        throw 'Could not determine the authenticated reviewer identity.'
    }
    $pr = Invoke-AzureDevOpsApi -Uri "$prUrl`?api-version=7.1" -Headers $headers
    if ($pr.status -ne 'active' -or $pr.isDraft) {
        throw 'PR voting requires an active, non-draft pull request.'
    }
    $sourceCommit = Get-BranchCommit -RefName $pr.sourceRefName
    $targetCommit = Get-BranchCommit -RefName $pr.targetRefName
    $reviewers = @($pr.reviewers | Where-Object { $_.id -eq $reviewerId.ToString() })
    $previousVote = 0
    if ($reviewers.Count -gt 1) {
        throw 'The authenticated reviewer occurs more than once.'
    }
    if ($reviewers.Count -eq 1) {
        if ($reviewers[0].isContainer -or $reviewers[0].vote -notin @(-10, -5, 0, 5, 10)) {
            throw 'The authenticated reviewer vote is invalid.'
        }
        $previousVote = [int]$reviewers[0].vote
    }

    if ($ReadState) {
        $snapshot = [ordered]@{
            reviewerId = $reviewerId.ToString()
            previousVote = $previousVote
            sourceCommit = $sourceCommit
            targetCommit = $targetCommit
        }
        [IO.File]::WriteAllText($OutputFile, (ConvertTo-Json -InputObject $snapshot -Compress), [Text.UTF8Encoding]::new($false))
        Write-Host "Current pull request vote: $(Get-VoteLabel $previousVote)"
        exit 0
    }

    if ($reviewerId -ne $ExpectedReviewerId -or $previousVote -ne $ExpectedVote) {
        throw 'PR vote unchanged: the reviewer identity or previous vote changed during the review.'
    }
    if ($sourceCommit -ine $ExpectedSourceCommit -or $targetCommit -ine $ExpectedTargetCommit) {
        throw 'PR vote unchanged: the source or target branch changed during the review. Run the task again.'
    }
    if ($previousVote -eq $Vote -and $Vote -lt 0) {
        Write-Host "Pull request vote unchanged: $(Get-VoteLabel $Vote)"
        exit 0
    }
    $voteBody = @{ id = $reviewerId.ToString(); vote = $Vote }
    if ($Vote -gt 0) {
        # Branch policies may require approval of the latest iteration even
        # when the account's displayed vote has not changed.
        $voteBody.isReapprove = $true
    }
    $body = ConvertTo-Json -Compress -InputObject $voteBody
    $updated = Invoke-AzureDevOpsApi -Uri "$prUrl/reviewers/$reviewerId`?api-version=7.1" -Headers $headers -Method Put -Body $body
    if ($updated.id -ne $reviewerId.ToString() -or $updated.vote -ne $Vote) {
        throw 'The PR voting API did not confirm the requested vote.'
    }
    if ($previousVote -eq $Vote) {
        Write-Host "Pull request vote confirmed: $(Get-VoteLabel $Vote)"
    } else {
        Write-Host "Pull request vote updated: $(Get-VoteLabel $previousVote) -> $(Get-VoteLabel $Vote)"
    }
} catch {
    Write-Host "##[error]$($_.Exception.Message)"
    exit 1
}

#endregion
