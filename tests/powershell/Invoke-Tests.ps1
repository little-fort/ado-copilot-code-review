<#
.SYNOPSIS
    Runs the PowerShell test suite with CI-appropriate settings.

.DESCRIPTION
    Thin wrapper around Invoke-Pester so local runs (npm run test:ps) and the
    CI workflow share identical configuration. Run.Exit makes the process exit
    non-zero when any test fails. Requires Pester 5 or later.
#>

$pester = Get-Module -ListAvailable Pester | Where-Object { $_.Version -ge [version]'5.0.0' }
if (-not $pester) {
    Write-Host "##[error]Pester 5+ is required. Install it with: Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser -SkipPublisherCheck"
    exit 1
}

$config = New-PesterConfiguration
$config.Run.Path = $PSScriptRoot
$config.Run.Exit = $true
$config.Output.Verbosity = 'Detailed'

Invoke-Pester -Configuration $config
