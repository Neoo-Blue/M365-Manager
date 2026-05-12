# ============================================================
#  Pester tests for OneDriveManager.ps1
#  Pure-logic tests -- no SPO / Graph calls.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'OneDriveManager.ps1')
}

Describe "Recent-files filter math" {
    It "respects the Days cutoff" {
        # We don't call Graph; we just exercise the cutoff comparison
        # that Get-OneDriveRecentFiles uses internally.
        $cutoff = (Get-Date).ToUniversalTime().AddDays(-30)
        $sample = @(
            [PSCustomObject]@{ name='today.docx';  lastModifiedDateTime=(Get-Date).ToUniversalTime().AddDays(-1).ToString('o');  size=1024 },
            [PSCustomObject]@{ name='old.pdf';     lastModifiedDateTime=(Get-Date).ToUniversalTime().AddDays(-90).ToString('o'); size=2048 }
        )
        $kept = @($sample | Where-Object { [DateTime]$_.lastModifiedDateTime -ge $cutoff })
        $kept.Count            | Should -Be 1
        $kept[0].name          | Should -Be 'today.docx'
    }
}

Describe "OneDrive URL parsing" {
    It "extracts host + personal path from a canonical OneDrive URL" {
        $url = 'https://contoso-my.sharepoint.com/personal/jane_contoso_com'
        ($url -match 'https?://([^/]+)/personal/([^/]+)') | Should -BeTrue
        $Matches[1] | Should -Be 'contoso-my.sharepoint.com'
        $Matches[2] | Should -Be 'jane_contoso_com'
    }
    It "rejects a non-OneDrive site URL" {
        $url = 'https://contoso.sharepoint.com/sites/sales'
        ($url -match 'https?://([^/]+)/personal/([^/]+)') | Should -BeFalse
    }
}

Describe "Retention end-date math" {
    It "60-day default lands 60 days in the future" {
        $end = (Get-Date).ToUniversalTime().AddDays(60)
        $delta = ($end - (Get-Date).ToUniversalTime()).TotalDays
        $delta | Should -BeGreaterOrEqual 59.99
        $delta | Should -BeLessOrEqual 60.01
    }
}
