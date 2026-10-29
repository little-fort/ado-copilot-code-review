<#
    Tests for pipeline logging-command neutralization (issues #47, #69).

    The Azure Pipelines agent acts on "##vso[...]" found ANYWHERE in a stdout
    or stderr line, not only at the start. Every script that echoes external
    text (PR content, work items, API responses) defines an identical
    ConvertTo-SafeLogText helper. These tests pin the helper's behavior, keep
    the copies from drifting apart, and run Get-AzureDevOpsPR.ps1 end to end
    to prove that indented thread text reaches the log neutralized while the
    context file handed to the review agent keeps the original text.

    Requires Pester 5+. Run via: pwsh -File tests/powershell/Invoke-Tests.ps1
#>

BeforeDiscovery {
    # Scripts that print external text and must define the helper
    $script:HelperScripts = @(
        'Add-AzureDevOpsPRComment.ps1',
        'Get-AzureDevOpsPR.ps1',
        'Get-AzureDevOpsPRChanges.ps1',
        'Get-AzureDevOpsWorkItems.ps1',
        'Get-JiraWorkItems.ps1'
    )
}

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:ScriptsDir = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts'

    # Returns the source text of a script's ConvertTo-SafeLogText function
    # (line endings normalized), or $null if the script does not define it
    function Get-HelperText {
        param([string]$ScriptName)

        $path = Join-Path $ScriptsDir $ScriptName
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            throw "$ScriptName failed to parse: $($parseErrors -join '; ')"
        }
        $funcAst = $ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'ConvertTo-SafeLogText'
        }, $true)
        if (-not $funcAst) {
            return $null
        }
        return ($funcAst.Extent.Text -replace "`r`n", "`n")
    }

    # Load the helper for direct unit testing
    . ([scriptblock]::Create((Get-HelperText 'Get-AzureDevOpsPR.ps1')))

    # Shared mock-API harness (HttpListener + child pwsh process)
    . (Join-Path $PSScriptRoot 'MockApiHelper.ps1')
}

Describe 'ConvertTo-SafeLogText definitions' {

    It 'is defined identically in <_>' -ForEach $HelperScripts {
        $text = Get-HelperText $_
        $text | Should -Not -BeNullOrEmpty -Because "$_ prints external text and must sanitize it"
        $text | Should -BeExactly (Get-HelperText 'Get-AzureDevOpsPR.ps1')
    }
}

Describe 'ConvertTo-SafeLogText behavior' {

    It 'breaks a logging command in the middle of a line' {
        ConvertTo-SafeLogText "Write-Host '##vso[task.logissue type=error]boom'" |
            Should -BeExactly "Write-Host '# #vso[task.logissue type=error]boom'"
    }

    It 'breaks an indented logging command' {
        # The original #47 fix only matched "##" at column 0, so indented
        # thread text (written with leading spaces) was never sanitized
        ConvertTo-SafeLogText '    ##vso[task.complete result=Failed;]done' |
            Should -BeExactly '    # #vso[task.complete result=Failed;]done'
    }

    It 'breaks formatting markers anywhere in the text' {
        ConvertTo-SafeLogText 'note ##[error]spoofed' | Should -BeExactly 'note # #[error]spoofed'
    }

    It 'breaks every occurrence across multiple lines' {
        $result = ConvertTo-SafeLogText "a ##vso[task.setvariable variable=X]1`n  b ##vso[task.prependpath]/tmp"
        $result | Should -Not -Match '##vso\['
        ([regex]::Matches($result, '# #vso\[')).Count | Should -Be 2
    }

    It 'matches the marker case-insensitively' {
        ConvertTo-SafeLogText 'x ##VSO[task.complete]' | Should -BeExactly 'x # #VSO[task.complete]'
    }

    It 'never leaves a marker behind in a run of hash characters' {
        ConvertTo-SafeLogText '###vso[a] ####vso[b]' | Should -Not -Match '##vso\['
    }

    It 'leaves ordinary text unchanged' {
        $text = "## Heading`nC# ## [link](url) #42"
        ConvertTo-SafeLogText $text | Should -BeExactly $text
    }

    It 'returns an empty string for null input' {
        ConvertTo-SafeLogText $null | Should -BeExactly ''
    }
}

Describe 'Get-AzureDevOpsPR.ps1 log output (issue #69)' {

    It 'neutralizes logging commands in PR and thread text on the console while the context file keeps the original' {
        $outDir = Join-Path $TestDrive 'pr-details'
        New-Item -ItemType Directory -Path $outDir | Out-Null
        $outFile = Join-Path $outDir 'PR_Details.txt'

        $prJson = @{
            pullRequestId = 7
            title         = 'Fix ##vso[task.setvariable variable=Leak]x'
            description   = "First line`n    ##vso[task.complete result=Failed;]done"
            status        = 'active'
            repository    = @{ name = 'TestRepo' }
            sourceRefName = 'refs/heads/feature'
            targetRefName = 'refs/heads/main'
            isDraft       = $false
            mergeStatus   = 'succeeded'
            createdBy     = @{ displayName = 'Dev'; uniqueName = 'dev@example.com' }
            creationDate  = '2026-01-01T00:00:00Z'
            reviewers     = @()
        } | ConvertTo-Json -Depth 5

        # Thread 11 is a human suggestion quoting a logging command (the #69
        # report); thread 12 is agent-authored, so it also lands in the JSON
        # section, which is printed as one long line
        $threadsJson = @{
            count = 2
            value = @(
                @{
                    id       = 11
                    status   = 'active'
                    comments = @(@{
                        id            = 1
                        commentType   = 'text'
                        content       = "Suggested fix:`n    Write-Host '##vso[task.logissue type=error]some message'"
                        author        = @{ displayName = 'Jane Reviewer'; uniqueName = 'jane@example.com' }
                        publishedDate = '2026-01-01T00:00:00Z'
                    })
                },
                @{
                    id       = 12
                    status   = 'active'
                    comments = @(@{
                        id            = 2
                        commentType   = 'text'
                        content       = "Avoid ##vso[task.setvariable] here.`n`n[Generated by GitHub Copilot]"
                        author        = @{ displayName = 'Jane Reviewer'; uniqueName = 'jane@example.com' }
                        publishedDate = '2026-01-01T00:00:00Z'
                    })
                }
            )
        } | ConvertTo-Json -Depth 6

        $prScript = Join-Path $ScriptsDir 'Get-AzureDevOpsPR.ps1'
        $result = Invoke-ScriptWithMockApi -Responses @(
            @{ Status = 200; Body = $prJson },
            @{ Status = 200; Body = '{"count":0,"value":[]}' },
            @{ Status = 200; Body = '{"count":1,"value":[{"id":1,"updatedDate":"2026-01-01T00:00:00Z"}]}' },
            @{ Status = 200; Body = $threadsJson }
        ) -BuildArgs {
            param($baseUrl)
            @('-File', $prScript,
              '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
              '-Project', 'TestProject', '-Repository', 'TestRepo',
              '-Id', '7', '-OutputFile', $outFile)
        }.GetNewClosure()

        $result.ExitCode | Should -Be 0
        $result.Output | Should -Not -Match '##vso\['
        $result.Errors | Should -Not -Match '##vso\['
        # Each source of external text was printed, just neutralized
        $result.Output | Should -Match ([regex]::Escape("Write-Host '# #vso[task.logissue type=error]some message'"))
        $result.Output | Should -Match ([regex]::Escape('Fix # #vso[task.setvariable variable=Leak]x'))
        $result.Output | Should -Match ([regex]::Escape('# #vso[task.complete result=Failed;]done'))
        $result.Output | Should -Match ([regex]::Escape('Avoid # #vso[task.setvariable] here.'))

        # The review agent's context file keeps the original text
        $fileText = Get-Content -Path $outFile -Raw
        $fileText | Should -Match ([regex]::Escape("Write-Host '##vso[task.logissue type=error]some message'"))
        $fileText | Should -Match ([regex]::Escape('Avoid ##vso[task.setvariable] here.'))
    }
}
