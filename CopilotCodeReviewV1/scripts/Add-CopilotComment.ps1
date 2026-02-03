<#
.SYNOPSIS
    Posts a comment to a pull request in Azure DevOps.

.DESCRIPTION
    This script is used by GitHub Copilot to add a comment to a pull request.
    It simplifies the calling process by populating the necessary parameters automatically
    from environment variables set by the pipeline task. Supports both general PR-level
    comments and file-specific inline comments.

.PARAMETER Comment
    Required. The comment text to post. Supports markdown formatting.

.PARAMETER Status
    Optional. Status for the new thread: 'Active' or 'Closed'. Default is 'Active'.

.PARAMETER FilePath
    Optional. File path for inline comment (e.g., '/src/MyProject/Program.cs').
    When provided with StartLine, creates an inline comment on the specified file.

.PARAMETER StartLine
    Optional. Starting line number for inline comment. Required when FilePath is provided.

.PARAMETER EndLine
    Optional. Ending line number for inline comment. Defaults to StartLine if not provided.

.EXAMPLE
    .\Add-CopilotComment.ps1 -Comment "This looks good!" -Status 'Closed'
    Creates a new general comment thread with closed status.

.EXAMPLE
    .\Add-CopilotComment.ps1 -Comment "Consider using async here" -Status 'Active' -FilePath '/src/Program.cs' -StartLine 42
    Creates an inline comment on line 42 of the specified file.

.EXAMPLE
    .\Add-CopilotComment.ps1 -Comment "This block needs refactoring" -Status 'Active' -FilePath '/src/Program.cs' -StartLine 42 -EndLine 50
    Creates an inline comment spanning lines 42-50 of the specified file.

.NOTES
    Author: Little Fort Software
    Date: December 2025
    Requires: PowerShell 5.1 or later
    
    Environment Variables Used:
    - AZUREDEVOPS_TOKEN: Authentication token (PAT or OAuth)
    - AZUREDEVOPS_AUTH_TYPE: 'Basic' for PAT, 'Bearer' for OAuth
    - ORGANIZATION: Azure DevOps organization name
    - PROJECT: Azure DevOps project name
    - REPOSITORY: Repository name
    - PRID: Pull request ID
    - ITERATION_ID: (Optional) PR iteration ID for inline comments
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Comment text to post")]
    [ValidateNotNullOrEmpty()]
    [string]$Comment,

    [Parameter(Mandatory = $false, HelpMessage = "Status for the new thread: Active or Closed")]
    [ValidateSet("Active", "Closed")]
    [string]$Status = 'Active',

    [Parameter(Mandatory = $false, HelpMessage = "File path for inline comment (e.g., '/src/MyProject/Program.cs')")]
    [string]$FilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Starting line number for inline comment")]
    [int]$StartLine,

    [Parameter(Mandatory = $false, HelpMessage = "Ending line number for inline comment (defaults to StartLine)")]
    [int]$EndLine
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Use the provided Status parameter (default: Active). The prompt should pass -Status when possible.
Write-Host "Posting comment with thread status: $Status" -ForegroundColor DarkGray

# Build the base parameters
$params = @{
    Token        = ${env:AZUREDEVOPS_TOKEN}
    AuthType     = ${env:AZUREDEVOPS_AUTH_TYPE}
    Organization = ${env:ORGANIZATION}
    Project      = ${env:PROJECT}
    Repository   = ${env:REPOSITORY}
    Id           = ${env:PRID}
    Comment      = $Comment
    Status       = $Status
}

# Add inline comment parameters if file path is provided
if ($FilePath) {
    $params.FilePath = $FilePath
    
    if ($StartLine -gt 0) {
        $params.StartLine = $StartLine
        
        # Default EndLine to StartLine if not provided
        if ($EndLine -gt 0) {
            $params.EndLine = $EndLine
        } else {
            $params.EndLine = $StartLine
        }
    }
    
    # Add iteration ID if available
    if (${env:ITERATION_ID}) {
        $params.IterationId = [int]${env:ITERATION_ID}
    }
    
    Write-Host "Posting inline comment on file: $FilePath (Lines $($params.StartLine)-$($params.EndLine))" -ForegroundColor DarkGray
}

& "$scriptDir\Add-AzureDevOpsPRComment.ps1" @params
