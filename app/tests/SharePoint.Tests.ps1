# ============================================================
#  Pester tests for SharePoint.ps1
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'Templates.ps1')
    . (Join-Path $script:RepoRoot 'SharePoint.ps1')
}

Describe "Stale-sites filter" {
    It "drops sites with LastContentModifiedDate inside the window" {
        $cutoff = (Get-Date).AddDays(-365)
        $sample = @(
            [PSCustomObject]@{ Url='https://x/sites/recent'; LastContentModifiedDate=(Get-Date).AddDays(-30) },
            [PSCustomObject]@{ Url='https://x/sites/old';    LastContentModifiedDate=(Get-Date).AddDays(-400) }
        )
        $stale = @($sample | Where-Object { $_.LastContentModifiedDate -lt $cutoff })
        $stale.Count   | Should -Be 1
        $stale[0].Url  | Should -Be 'https://x/sites/old'
    }
}

Describe "System-account filter on Get-SiteOwners (predicate)" {
    It "filters out the well-known service-principal shapes" {
        $owners = @(
            [PSCustomObject]@{ LoginName='c:0t.c|tenant|adminagents'; DisplayName='SP Admins' },
            [PSCustomObject]@{ LoginName='SHAREPOINT\system';         DisplayName='SHAREPOINT\system' },
            [PSCustomObject]@{ LoginName='app@sharepoint';            DisplayName='App SP' },
            [PSCustomObject]@{ LoginName='jane@contoso.com';          DisplayName='Jane' }
        )
        $human = @($owners | Where-Object { $_.LoginName -notmatch '^(c:0|SHAREPOINT\\|app@sharepoint|i:0#\.f\|membership\|app_)' })
        $human.Count | Should -Be 1
        $human[0].LoginName | Should -Be 'jane@contoso.com'
    }
}

Describe "Site template JSON shapes" {
    It "site-project.json + site-team.json parse" {
        $project = Get-SiteTemplate -Name 'project'
        $project.alias    | Should -Not -BeNullOrEmpty
        $project.template | Should -Not -BeNullOrEmpty
        $team = Get-SiteTemplate -Name 'team'
        $team.alias    | Should -Not -BeNullOrEmpty
        $team.template | Should -Not -BeNullOrEmpty
    }
}

Describe "Owner add/remove reverse-recipe pair" {
    It "AddSiteOwner has a matching RemoveSiteOwner inverse target shape" {
        # Both should accept {siteUrl, userUpn}. We can't run the
        # cmdlet but we can sanity check the parameter expectations
        # by inspecting the handler scriptblocks if Undo.ps1 has
        # been dot-sourced.
        . (Join-Path $script:RepoRoot 'Undo.ps1')
        $script:UndoHandlers['AddSiteOwner']    | Should -Not -BeNullOrEmpty
        $script:UndoHandlers['RemoveSiteOwner'] | Should -Not -BeNullOrEmpty
    }
}
