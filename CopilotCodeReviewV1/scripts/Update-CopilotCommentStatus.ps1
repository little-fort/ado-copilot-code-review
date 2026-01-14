<#
.SYNOPSIS
    Updates the status of a comment thread in a pull request.

.DESCRIPTION
    This script is used by GitHub Copilot to update the status of a comment thread
    in a pull request. It simplifies the calling process by populating the necessary
    parameters automatically from environment variables set by the pipeline task.

.PARAMETER ThreadId
    Required. The ID of the comment thread to update.

.PARAMETER Status
    Required. The new status for the thread. Valid values: Active, Fixed, WontFix, Closed, Pending.

.EXAMPLE
    .\Update-CopilotCommentStatus.ps1 -ThreadId 456 -Status 'Fixed'
    Marks thread #456 as Fixed.

.EXAMPLE
    .\Update-CopilotCommentStatus.ps1 -ThreadId 789 -Status 'Closed'
    Closes thread #789.

.NOTES
    Author: Little Fort Software
    Date: January 2026
    Requires: PowerShell 5.1 or later
    
    Environment Variables Used:
    - AZUREDEVOPS_TOKEN: Authentication token (PAT or OAuth)
    - AZUREDEVOPS_AUTH_TYPE: 'Basic' for PAT, 'Bearer' for OAuth
    - ORGANIZATION: Azure DevOps organization name
    - PROJECT: Azure DevOps project name
    - REPOSITORY: Repository name
    - PRID: Pull request ID
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Thread ID to update")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$ThreadId,

    [Parameter(Mandatory = $true, HelpMessage = "New status for the thread")]
    [ValidateSet("Active", "Fixed", "WontFix", "Closed", "Pending")]
    [string]$Status
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "Updating thread #$ThreadId to status: $Status" -ForegroundColor DarkGray

& "$scriptDir\Update-AzureDevOpsPRCommentStatus.ps1" `
    -Token ${env:AZUREDEVOPS_TOKEN} `
    -AuthType ${env:AZUREDEVOPS_AUTH_TYPE} `
    -Organization ${env:ORGANIZATION} `
    -Project ${env:PROJECT} `
    -Repository ${env:REPOSITORY} `
    -Id ${env:PRID} `
    -ThreadId $ThreadId `
    -Status $Status
