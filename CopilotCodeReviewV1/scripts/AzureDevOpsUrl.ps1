<#
.SYNOPSIS
    Shared helpers for building Azure DevOps REST and web base URLs.

.DESCRIPTION
    Centralizes URL construction for Azure DevOps Services and Server (on-prem).
    Uses $env:AZUREDEVOPS_ONPREMISE to determine whether to use on-prem URLs.
#>

function Get-AzureDevOpsBaseUrls {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Project,

        [Parameter(Mandatory = $false)]
        [string]$Organization,

        [Parameter(Mandatory = $false)]
        [switch]$Silent
    )

    $onPremise = $false
    if ($env:AZUREDEVOPS_ONPREMISE -match '^(?i:true|1|yes)$') {
        $onPremise = $true
    }

    $collectionUri = $null
    if ($onPremise) {
        $collectionUri = $env:SYSTEM_COLLECTIONURI
        if ([string]::IsNullOrWhiteSpace($collectionUri)) {
            if (-not $Silent) {
                Write-Error "SYSTEM_COLLECTIONURI is required when On-Premise mode is enabled."
            }
            return $null
        }
        $collectionUri = $collectionUri.Trim()
        if (-not $collectionUri.EndsWith('/')) {
            $collectionUri += '/'
        }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($Organization)) {
            if (-not $Silent) {
                Write-Error "Organization is required when On-Premise mode is disabled."
            }
            return $null
        }
    }

    $apiBaseUrl = if ($onPremise) { "$collectionUri$Project/_apis" } else { "https://dev.azure.com/$Organization/$Project/_apis" }
    $webBaseUrl = if ($onPremise) { "$collectionUri$Project" } else { "https://dev.azure.com/$Organization/$Project" }

    return [PSCustomObject]@{
        ApiBaseUrl   = $apiBaseUrl
        WebBaseUrl   = $webBaseUrl
        OnPremise    = $onPremise
        CollectionUri = $collectionUri
    }
}
