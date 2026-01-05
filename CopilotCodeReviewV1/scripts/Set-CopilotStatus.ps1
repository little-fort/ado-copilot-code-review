<#
.SYNOPSIS
    Helper script for Copilot CLI to set PR status.

.DESCRIPTION
    This script is used by GitHub Copilot to set a status on a pull request.
    It simplifies the calling process by populating the necessary parameters automatically
    from environment variables set by the pipeline task.

.PARAMETER State
    Required. The state of the status check: succeeded, failed, error, pending, notSet, notApplicable

.PARAMETER Description
    Optional. A description of the status check.

.EXAMPLE
    .\Set-CopilotStatus.ps1 -State "succeeded" -Description "Code review passed"
    Sets a succeeded status.

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
    [Parameter(Mandatory = $true, HelpMessage = "Status state")]
    [ValidateSet("succeeded", "failed", "error", "pending", "notSet", "notApplicable")]
    [string]$State,

    [Parameter(Mandatory = $false, HelpMessage = "Description of the status")]
    [string]$Description = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "Setting PR status: $State" -ForegroundColor DarkGray

& "$scriptDir\Set-AzureDevOpsPRStatus.ps1" `
    -Token ${env:AZUREDEVOPS_TOKEN} `
    -AuthType ${env:AZUREDEVOPS_AUTH_TYPE} `
    -Organization ${env:ORGANIZATION} `
    -Project ${env:PROJECT} `
    -Repository ${env:REPOSITORY} `
    -Id ${env:PRID} `
    -State $State `
    -Description $Description
