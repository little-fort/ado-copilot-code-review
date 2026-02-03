<#
.SYNOPSIS
    Retrieves commits and changed files from the most recent iteration of a pull request.

.DESCRIPTION
    This script uses the Azure DevOps REST API to get the list of commits and 
    changed files from the most recent iteration (latest push) of a pull request.

.PARAMETER Token
    Required. Authentication token for Azure DevOps. Can be a PAT or OAuth token.

.PARAMETER AuthType
    Optional. The type of authentication to use. Valid values: 'Basic' (for PAT) or 'Bearer' (for OAuth/System.AccessToken).
    Default is 'Basic'.

.PARAMETER Organization
    Required. The Azure DevOps organization name.

.PARAMETER Project
    Required. The Azure DevOps project name.

.PARAMETER Repository
    Required. The repository name where the pull request exists.

.PARAMETER Id
    Required. The pull request ID to retrieve changes for.

.EXAMPLE
    .\Get-AzureDevOpsPRChanges.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123
    Retrieves the commits and changed files from the most recent iteration of PR #123.

.EXAMPLE
    .\Get-AzureDevOpsPRChanges.ps1 -Token "oauth-token" -AuthType "Bearer" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123
    Retrieves PR changes using OAuth/System.AccessToken authentication.

.EXAMPLE
    .\Get-AzureDevOpsPRChanges.ps1 -Token "your-pat" -Organization "myorg" -Project "myproject" -Repository "myrepo" -Id 123 -OutputFile "C:\output\pr-changes.txt"
    Writes the pull request changes to the specified file.

.NOTES
    Author: Little Fort Software
    Date: December 2025
    Requires: PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Authentication token for Azure DevOps (PAT or OAuth token)")]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    [Parameter(Mandatory = $false, HelpMessage = "Authentication type: 'Basic' for PAT, 'Bearer' for OAuth")]
    [ValidateSet("Basic", "Bearer")]
    [string]$AuthType = "Basic",

    [Parameter(Mandatory = $true, HelpMessage = "Azure DevOps organization name")]
    [ValidateNotNullOrEmpty()]
    [string]$Organization,

    [Parameter(Mandatory = $true, HelpMessage = "Azure DevOps project name")]
    [ValidateNotNullOrEmpty()]
    [string]$Project,

    [Parameter(Mandatory = $true, HelpMessage = "Repository name")]
    [ValidateNotNullOrEmpty()]
    [string]$Repository,

    [Parameter(Mandatory = $true, HelpMessage = "Pull request ID")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$Id,

    [Parameter(Mandatory = $false, HelpMessage = "Output file path to write results to")]
    [string]$OutputFile
)

#region Helper Functions

function Write-Output-Line {
    param(
        [string]$Message = "",
        [string]$ForegroundColor = "White",
        [switch]$NoNewline
    )
    
    if ($script:OutputToFile) {
        if ($NoNewline) {
            $script:OutputBuilder.Append($Message) | Out-Null
        }
        else {
            $script:OutputBuilder.AppendLine($Message) | Out-Null
        }
    }
    
    if ($NoNewline) {
        Write-Host $Message -ForegroundColor $ForegroundColor -NoNewline
    }
    else {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
}

function Get-AuthorizationHeader {
    param(
        [string]$Token,
        [string]$AuthType = "Basic"
    )
    
    if ($AuthType -eq "Bearer") {
        return @{
            Authorization  = "Bearer $Token"
            "Content-Type" = "application/json"
        }
    }
    else {
        $base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Token"))
        return @{
            Authorization  = "Basic $base64Auth"
            "Content-Type" = "application/json"
        }
    }
}

function Invoke-AzureDevOpsApi {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "Get"
    )
    
    try {
        $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ErrorAction Stop
        return $response
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $errorMessage = $_.ErrorDetails.Message
        
        if ($statusCode -eq 401) {
            Write-Error "Authentication failed. Please verify your PAT is valid and has appropriate permissions."
        }
        elseif ($statusCode -eq 404) {
            Write-Error "Resource not found. Please verify the organization, project, repository, and PR ID."
        }
        else {
            Write-Error "API request failed: $errorMessage (Status: $statusCode)"
        }
        return $null
    }
}

function Format-DateForDisplay {
    param([string]$DateString)
    
    if ([string]::IsNullOrEmpty($DateString)) {
        return "N/A"
    }
    
    try {
        $date = [DateTime]::Parse($DateString)
        return $date.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return $DateString
    }
}

function Get-ChangeTypeDisplay {
    param([string]$ChangeType)
    
    switch ($ChangeType) {
        "add"      { return @{ Text = "Added"; Color = "Green" } }
        "edit"     { return @{ Text = "Modified"; Color = "Yellow" } }
        "delete"   { return @{ Text = "Deleted"; Color = "Red" } }
        "rename"   { return @{ Text = "Renamed"; Color = "Cyan" } }
        "copy"     { return @{ Text = "Copied"; Color = "Cyan" } }
        default    { return @{ Text = $ChangeType; Color = "White" } }
    }
}

#endregion

#region Main Logic

# Initialize output handling
$script:OutputToFile = -not [string]::IsNullOrEmpty($OutputFile)
$script:OutputBuilder = [System.Text.StringBuilder]::new()

$headers = Get-AuthorizationHeader -Token $Token -AuthType $AuthType
$baseUrl = "https://dev.azure.com/$Organization/$Project/_apis/git/repositories/$Repository/pullrequests/$Id"
$apiVersion = "api-version=7.1"

# Verify the PR exists
Write-Host "`nRetrieving pull request #$Id..." -ForegroundColor Cyan
$prUrl = "$baseUrl`?$apiVersion"
$pr = Invoke-AzureDevOpsApi -Uri $prUrl -Headers $headers

if ($null -eq $pr) {
    exit 1
}

Write-Host "Found PR: $($pr.title)" -ForegroundColor Green
Write-Host "Status: $($pr.status.ToUpper())" -ForegroundColor $(if ($pr.status -eq "active") { "Green" } else { "Yellow" })

# Get iterations
Write-Host "`nRetrieving iterations..." -ForegroundColor Cyan
$iterationsUrl = "$baseUrl/iterations?$apiVersion"
$iterations = Invoke-AzureDevOpsApi -Uri $iterationsUrl -Headers $headers

if ($null -eq $iterations -or $iterations.count -eq 0) {
    Write-Warning "No iterations found for this pull request."
    exit 0
}

$latestIteration = $iterations.value | Sort-Object -Property id -Descending | Select-Object -First 1
$iterationId = $latestIteration.id

Write-Host "Found $($iterations.count) iteration(s). Using latest: Iteration #$iterationId" -ForegroundColor Green

# Get commits for the PR
Write-Host "`nRetrieving commits..." -ForegroundColor Cyan
$commitsUrl = "$baseUrl/commits?$apiVersion"
$commits = Invoke-AzureDevOpsApi -Uri $commitsUrl -Headers $headers

# Get changes for the latest iteration
Write-Host "Retrieving changes for iteration #$iterationId..." -ForegroundColor Cyan
$changesUrl = "$baseUrl/iterations/$iterationId/changes?$apiVersion"
$changes = Invoke-AzureDevOpsApi -Uri $changesUrl -Headers $headers

# Display results
Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray
Write-Output-Line "PULL REQUEST CHANGES - ITERATION #$iterationId" -ForegroundColor Green
Write-Output-Line ("=" * 80) -ForegroundColor DarkGray

# Iteration Info
Write-Output-Line "`n[Iteration Details]" -ForegroundColor Yellow
Write-Output-Line "  Iteration ID:     #$iterationId"
Write-Output-Line "  Created:          $(Format-DateForDisplay $latestIteration.createdDate)"
Write-Output-Line "  Updated:          $(Format-DateForDisplay $latestIteration.updatedDate)"
if ($latestIteration.sourceRefCommit) {
    Write-Output-Line "  Source Commit:    $($latestIteration.sourceRefCommit.commitId.Substring(0, 8))"
}
if ($latestIteration.targetRefCommit) {
    Write-Output-Line "  Target Commit:    $($latestIteration.targetRefCommit.commitId.Substring(0, 8))"
}

# Commits
Write-Output-Line "`n[Commits in this PR]" -ForegroundColor Yellow
if ($commits -and $commits.value -and $commits.value.Count -gt 0) {
    Write-Output-Line "  Total commits: $($commits.value.Count)`n"
    
    foreach ($commit in $commits.value) {
        $shortId = $commit.commitId.Substring(0, 8)
        $message = $commit.comment -split "`n" | Select-Object -First 1
        if ($message.Length -gt 60) {
            $message = $message.Substring(0, 57) + "..."
        }
        Write-Output-Line "  $shortId - $message" -ForegroundColor Cyan
        Write-Output-Line "           Author: $($commit.author.name) | $(Format-DateForDisplay $commit.author.date)" -ForegroundColor DarkGray
    }
}
else {
    Write-Output-Line "  No commits found."
}

# Changed Files
Write-Output-Line "`n[Changed Files]" -ForegroundColor Yellow
if ($changes -and $changes.changeEntries -and $changes.changeEntries.Count -gt 0) {
    # Group by change type for summary
    $addedCount = ($changes.changeEntries | Where-Object { $_.changeType -eq "add" }).Count
    $modifiedCount = ($changes.changeEntries | Where-Object { $_.changeType -eq "edit" }).Count
    $deletedCount = ($changes.changeEntries | Where-Object { $_.changeType -eq "delete" }).Count
    $otherCount = $changes.changeEntries.Count - $addedCount - $modifiedCount - $deletedCount
    
    Write-Output-Line "  Total files changed: $($changes.changeEntries.Count)"
    $summaryLine = "  +$addedCount added | ~$modifiedCount modified | -$deletedCount deleted"
    if ($otherCount -gt 0) {
        $summaryLine += " | $otherCount other"
    }
    Write-Output-Line $summaryLine
    Write-Output-Line ""
    
    # List each file
    foreach ($change in $changes.changeEntries) {
        $changeDisplay = Get-ChangeTypeDisplay -ChangeType $change.changeType
        $filePath = $change.item.path
        
        Write-Output-Line "  [$($changeDisplay.Text)] $filePath" -ForegroundColor $changeDisplay.Color
        
        # Show original path for renames
        if ($change.changeType -eq "rename" -and $change.originalPath) {
            Write-Output-Line "         (from: $($change.originalPath))" -ForegroundColor DarkGray
        }
    }
}
else {
    Write-Output-Line "  No file changes found in this iteration."
}

Write-Output-Line ("`n" + ("=" * 80)) -ForegroundColor DarkGray

# Provide link to the PR
$webUrl = "https://dev.azure.com/$Organization/$Project/_git/$Repository/pullrequest/$Id"
Write-Host "`nView PR: $webUrl" -ForegroundColor Cyan
if ($script:OutputToFile) {
    $script:OutputBuilder.AppendLine("`nView PR: $webUrl") | Out-Null
}

# Write to output file if specified
if ($script:OutputToFile) {
    try {
        $outputDir = Split-Path -Parent $OutputFile
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        $script:OutputBuilder.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "`nOutput written to: $OutputFile" -ForegroundColor Green
        
        # Also write the iteration ID to a separate file for use by other scripts
        $iterationIdFile = Join-Path $outputDir "Iteration_Id.txt"
        $iterationId.ToString() | Out-File -FilePath $iterationIdFile -Encoding UTF8 -NoNewline
        Write-Host "Iteration ID written to: $iterationIdFile" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to write output file: $_"
    }
}

#endregion
