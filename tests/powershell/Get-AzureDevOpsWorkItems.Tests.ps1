<#
    Tests for Get-AzureDevOpsWorkItems.ps1: unit tests for the ConvertFrom-Html
    helper, plus batch-fetch tests that run the script against a mock API
    (MockApiHelper.ps1).

    The function is extracted from the script via the PowerShell AST rather than
    dot-sourcing the whole file (which would execute its main logic and demand
    mandatory parameters). This keeps the scripts standalone, per repo convention.

    Requires Pester 5+. Run via: pwsh -File tests/powershell/Invoke-Tests.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $scriptPath = Join-Path $repoRoot 'CopilotCodeReviewV1/scripts/Get-AzureDevOpsWorkItems.ps1'

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw "Get-AzureDevOpsWorkItems.ps1 failed to parse: $($parseErrors -join '; ')"
    }

    $funcAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'ConvertFrom-Html'
    }, $true)
    if (-not $funcAst) {
        throw 'ConvertFrom-Html function not found in Get-AzureDevOpsWorkItems.ps1 — update this test if it was renamed or moved.'
    }

    # Bring just the function under test into scope
    . ([scriptblock]::Create($funcAst.Extent.Text))
}

Describe 'ConvertFrom-Html' {

    Context 'strikethrough preservation (issue #52)' {

        It 'converts del tags to ~~text~~' {
            ConvertFrom-Html '<del>old criteria</del>' | Should -Be '~~old criteria~~'
        }

        It 'converts strike tags to ~~text~~' {
            ConvertFrom-Html '<strike>retracted</strike>' | Should -Be '~~retracted~~'
        }

        It 'converts s tags to ~~text~~' {
            ConvertFrom-Html '<s>superseded</s>' | Should -Be '~~superseded~~'
        }

        It 'is case-insensitive' {
            ConvertFrom-Html '<DEL>old</DEL>' | Should -Be '~~old~~'
        }

        It 'handles attributes on the strikethrough tag' {
            ConvertFrom-Html '<del style="color:red" data-x="1">old</del>' | Should -Be '~~old~~'
        }

        It 'handles multi-line struck-through content' {
            $result = ConvertFrom-Html "<del>first line`nsecond line</del>"
            $result | Should -Match '^~~first line'
            $result | Should -Match 'second line~~$'
        }

        It 'strips inner formatting tags but keeps the strikethrough markers' {
            ConvertFrom-Html '<del><b>bold old text</b></del>' | Should -Be '~~bold old text~~'
        }

        It 'keeps struck-through and live criteria distinguishable' {
            $html = '<div>Must support login</div><div><del>Must support SSO</del></div><div>Must log errors</div>'
            $result = ConvertFrom-Html $html
            $result | Should -Match 'Must support login'
            $result | Should -Match '~~Must support SSO~~'
            $result | Should -Match 'Must log errors'
        }

        It 'does not treat span or strong tags as strikethrough despite starting with s' {
            ConvertFrom-Html '<span>plain</span>' | Should -Be 'plain'
            ConvertFrom-Html '<strong>important</strong>' | Should -Be 'important'
        }

        It 'strips an unclosed del tag rather than corrupting output' {
            ConvertFrom-Html '<del>dangling text' | Should -Be 'dangling text'
        }
    }

    Context 'existing conversion behavior (regression)' {

        It 'converts br tags to newlines' {
            ConvertFrom-Html 'line one<br>line two' | Should -Be "line one`nline two"
        }

        It 'converts block-level closing tags to newlines' {
            $result = ConvertFrom-Html '<p>para one</p><p>para two</p>'
            $result | Should -Match "para one`n"
            $result | Should -Match 'para two'
        }

        It 'decodes HTML entities' {
            ConvertFrom-Html 'a &amp; b &lt;c&gt;' | Should -Be 'a & b <c>'
        }

        It 'returns an empty string for null or whitespace input' {
            ConvertFrom-Html $null | Should -Be ''
            ConvertFrom-Html '   ' | Should -Be ''
        }

        It 'collapses runs of blank lines' {
            $result = ConvertFrom-Html "<p>a</p>`n`n`n`n<p>b</p>"
            $result | Should -Not -Match "`n{3,}"
        }
    }
}

Describe 'Batch fetch' {

    BeforeAll {
        . (Join-Path $PSScriptRoot 'MockApiHelper.ps1')

        # Runs the script for work items 101 and 102 with a review state folder;
        # returns the harness result plus the output file text and state folder
        function Invoke-WorkItemsScript {
            param([string]$ResponseBody)
            $stateDir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $stateDir | Out-Null
            $outFile = Join-Path $stateDir 'Work_Item_Details.txt'
            $workItemsScript = $scriptPath
            $result = Invoke-ScriptWithMockApi -Responses @(
                @{ Status = 200; Body = $ResponseBody }
            ) -Environment @{ COPILOT_REVIEW_STATE_DIR = $stateDir } -BuildArgs {
                param($baseUrl)
                @('-File', $workItemsScript, '-Token', 'test-token', '-CollectionUri', "$baseUrl/testorg",
                  '-Project', 'TestProject', '-WorkItemIds', '101,102', '-OutputFile', $outFile)
            }.GetNewClosure()
            $fileText = if (Test-Path $outFile) { Get-Content -Path $outFile -Raw } else { '' }
            $result |
                Add-Member -NotePropertyName FileText -NotePropertyValue $fileText -PassThru |
                Add-Member -NotePropertyName StateDir -NotePropertyValue $stateDir -PassThru
        }
    }

    # Azure DevOps fails the whole batch when one ID is deleted or unreadable,
    # unless errorPolicy=omit, which returns null in that item's place
    It 'skips an unreadable work item and names it, without a degradation marker' {
        $run = Invoke-WorkItemsScript -ResponseBody '{"count":2,"value":[{"id":101,"fields":{"System.WorkItemType":"Bug","System.Title":"Crash on save","System.State":"New"}},null]}'

        $run.ExitCode | Should -Be 0
        $run.Requests.Count | Should -Be 1
        $run.Requests[0].Path | Should -Match 'errorPolicy=omit'
        $run.FileText | Should -Match '\[Work Item #101 - Bug\]'
        $run.FileText | Should -Match 'Crash on save'
        $run.FileText | Should -Match 'Work item\(s\) #102 could not be read'
        $run.FileText | Should -Not -Match '\[Work Item #(102)? -'
        Test-Path (Join-Path $run.StateDir 'context') | Should -BeFalse
    }

    It 'names every work item and writes no degradation marker when none can be read' {
        $run = Invoke-WorkItemsScript -ResponseBody '{"count":2,"value":[null,null]}'

        $run.ExitCode | Should -Be 0
        $run.FileText | Should -Match 'Work item\(s\) #101, #102 could not be read'
        $run.FileText | Should -Not -Match '\[Work Item #'
        Test-Path (Join-Path $run.StateDir 'context') | Should -BeFalse
    }
}
